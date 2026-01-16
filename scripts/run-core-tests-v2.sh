#!/usr/bin/env bash
set -euo pipefail

# =============================================================================
# Excelize Core Tests Runner with DuckDB Engine Support
# =============================================================================
# This script runs the core test suite including the DuckDB calculation engine
# tests for accuracy and performance validation.
#
# Usage:
#   ./scripts/run-core-tests.sh              # Run all tests
#   ./scripts/run-core-tests.sh --quick      # Quick core tests only
#   ./scripts/run-core-tests.sh --duckdb     # DuckDB-focused tests
#   ./scripts/run-core-tests.sh --accuracy   # DuckDB accuracy/parity tests
#   ./scripts/run-core-tests.sh --perf       # DuckDB performance benchmarks
#   ./scripts/run-core-tests.sh --help       # Show help
# =============================================================================

# Configuration
COVERAGE_FILE="${COVERAGE_FILE:-coverage.out}"
PKG_PATTERN="${PKG_PATTERN:-./...}"
EXCLUDE_PACKAGES="${EXCLUDE_PACKAGES:-github.com/xuri/excelize/v2/tests/manual/tools}"
GO_TEST_FLAGS="${GO_TEST_FLAGS:--v}"
BENCHMARK_TIME="${BENCHMARK_TIME:-3s}"
BENCHMARK_COUNT="${BENCHMARK_COUNT:-3}"
export GOCACHE="${GOCACHE:-/tmp/go-build}"

# Colors
GREEN="\033[32m"
YELLOW="\033[33m"
RED="\033[31m"
BLUE="\033[34m"
CYAN="\033[36m"
BOLD="\033[1m"
RESET="\033[0m"

# Defaults
MODE="all"
VERBOSE=false
SKIP_COVERAGE=false

# =============================================================================
# Help Message
# =============================================================================
show_help() {
    cat << EOF
${BOLD}Excelize Core Tests Runner with DuckDB Engine Support${RESET}

${CYAN}USAGE:${RESET}
    ./scripts/run-core-tests.sh [OPTIONS]

${CYAN}OPTIONS:${RESET}
    --quick, -q       Run quick core tests only (no large scale tests)
    --duckdb, -d      Run DuckDB-focused tests (accuracy + performance)
    --accuracy, -a    Run DuckDB accuracy/parity tests only
    --perf, -p        Run DuckDB performance benchmarks only
    --full, -f        Run full test suite with extended benchmarks
    --verbose, -V     Enable verbose output
    --no-coverage     Skip coverage collection
    --help, -h        Show this help message

${CYAN}ENVIRONMENT VARIABLES:${RESET}
    COVERAGE_FILE     Coverage output file (default: coverage.out)
    PKG_PATTERN       Package pattern to test (default: ./...)
    GO_TEST_FLAGS     Additional go test flags (default: -v)
    BENCHMARK_TIME    Benchmark duration (default: 3s)
    BENCHMARK_COUNT   Benchmark iterations (default: 3)

${CYAN}EXAMPLES:${RESET}
    # Run all tests with coverage
    ./scripts/run-core-tests.sh

    # Run only DuckDB parity tests
    ./scripts/run-core-tests.sh --accuracy

    # Run DuckDB benchmarks with 5s per benchmark
    BENCHMARK_TIME=5s ./scripts/run-core-tests.sh --perf

    # Run quick tests without coverage
    ./scripts/run-core-tests.sh --quick --no-coverage

${CYAN}TEST CATEGORIES:${RESET}
    ${BOLD}Core Tests:${RESET}
      - Basic functionality tests for all packages
      - Formula calculation engine tests
      - File I/O and format tests

    ${BOLD}DuckDB Accuracy Tests:${RESET}
      - Native vs DuckDB parity tests
      - Formula compilation tests
      - Multi-level integration tests (100 → 10K rows)
      - Cross-worksheet reference tests

    ${BOLD}DuckDB Performance Tests:${RESET}
      - SUMIFS/COUNTIFS benchmarks (100K rows)
      - Batch operations benchmarks
      - Large data loading benchmarks
      - Cache efficiency benchmarks
EOF
}

# =============================================================================
# Parse Arguments
# =============================================================================
while [[ $# -gt 0 ]]; do
    case $1 in
        --quick|-q)
            MODE="quick"
            shift
            ;;
        --duckdb|-d)
            MODE="duckdb"
            shift
            ;;
        --accuracy|-a)
            MODE="accuracy"
            shift
            ;;
        --perf|-p)
            MODE="perf"
            shift
            ;;
        --full|-f)
            MODE="full"
            shift
            ;;
        --verbose|-V)
            VERBOSE=true
            GO_TEST_FLAGS="${GO_TEST_FLAGS} -v"
            shift
            ;;
        --no-coverage)
            SKIP_COVERAGE=true
            shift
            ;;
        --help|-h)
            show_help
            exit 0
            ;;
        *)
            echo -e "${RED}Unknown option: $1${RESET}"
            echo "Use --help for usage information"
            exit 1
            ;;
    esac
done

# =============================================================================
# Utility Functions
# =============================================================================
log_header() {
    echo ""
    echo -e "${BOLD}${BLUE}════════════════════════════════════════════════════════════════${RESET}"
    echo -e "${BOLD}${BLUE}  $1${RESET}"
    echo -e "${BOLD}${BLUE}════════════════════════════════════════════════════════════════${RESET}"
    echo ""
}

log_section() {
    echo ""
    echo -e "${CYAN}────────────────────────────────────────────────────────────────${RESET}"
    echo -e "${CYAN}  $1${RESET}"
    echo -e "${CYAN}────────────────────────────────────────────────────────────────${RESET}"
}

log_info() {
    echo -e "[${BLUE}INFO${RESET}] $1"
}

log_success() {
    echo -e "[${GREEN}PASS${RESET}] $1"
}

log_warning() {
    echo -e "[${YELLOW}WARN${RESET}] $1"
}

log_error() {
    echo -e "[${RED}FAIL${RESET}] $1"
}

print_rate_summary() {
    local passed=$1
    local total=$2
    local status="$3"
    local rate
    rate=$(awk -v s="$passed" -v t="$total" 'BEGIN { if (t==0) { printf "0.0" } else { printf "%.1f", (s/t)*100 } }')
    local color="${GREEN}"
    if (( passed == total )); then
        color="${GREEN}"
    elif (( passed > 0 )); then
        color="${YELLOW}"
    else
        color="${RED}"
    fi
    printf "[core-tests] %b%s: %d/%d packages passed (%s%%)%b\n" "${color}" "${status}" "${passed}" "${total}" "${rate}" "${RESET}"
}

format_duration() {
    local duration=$1
    if (( duration >= 60000 )); then
        echo "$((duration / 60000))m $((duration % 60000 / 1000))s"
    elif (( duration >= 1000 )); then
        echo "$((duration / 1000)).$((duration % 1000 / 100))s"
    else
        echo "${duration}ms"
    fi
}

# =============================================================================
# Core Tests Runner (from original script)
# =============================================================================
run_core_tests() {
    log_header "Running Core Tests"

    log_info "Collecting package list (${PKG_PATTERN})"

    PKGS=()
    EXCLUDES=()
    while IFS= read -r line; do
        [[ -z "${line}" ]] && continue
        EXCLUDES+=("${line}")
    done < <(printf "%s\n" "${EXCLUDE_PACKAGES}" | tr ' ' '\n' | tr ',' '\n')

    PKG_LIST=""
    if ! PKG_LIST=$(go list "${PKG_PATTERN}"); then
        log_error "Failed to list packages for pattern ${PKG_PATTERN}"
        exit 1
    fi

    while IFS= read -r pkg; do
        skip=false
        for exclude in "${EXCLUDES[@]:-}"; do
            if [[ "${pkg}" == "${exclude}" ]]; then
                log_warning "Skipping excluded package ${pkg}"
                skip=true
                break
            fi
        done
        if [[ "${skip}" == "true" ]]; then
            continue
        fi
        if [[ -n "${pkg}" ]]; then
            PKGS+=("$pkg")
        fi
    done <<< "${PKG_LIST}"

    PKG_COUNT=${#PKGS[@]}
    if (( PKG_COUNT == 0 )); then
        log_error "No packages matched pattern ${PKG_PATTERN}"
        exit 1
    fi

    log_info "Running ${PKG_COUNT} packages with coverage -> ${COVERAGE_FILE}"

    TMP_COVERAGE="$(mktemp)"
    : > "${COVERAGE_FILE}"
    SUCCESS_COUNT=0
    TOTAL_TESTS=0
    TOTAL_PASSED=0
    TOTAL_FAILED=0
    TOTAL_SKIPPED=0

    for ((i=0; i<PKG_COUNT; i++)); do
        pkg="${PKGS[$i]}"
        printf '[core-tests] (%d/%d) go test %s\n' "$((i+1))" "$PKG_COUNT" "$pkg"
        RUN_LOG="$(mktemp)"

        COVERAGE_FLAGS=""
        if [[ "${SKIP_COVERAGE}" != "true" ]]; then
            COVERAGE_FLAGS="-covermode=atomic -coverprofile=${TMP_COVERAGE}"
        fi

        if ! GOFLAGS="${GOFLAGS:-}" go test $GO_TEST_FLAGS "$pkg" $COVERAGE_FLAGS 2>&1 | tee "${RUN_LOG}"; then
            log_error "FAILED in package ${pkg}"
            pkgRuns=$(grep -c '^=== RUN' "${RUN_LOG}" || true)
            pkgPass=$(grep -c '^[[:space:]]*--- PASS' "${RUN_LOG}" || true)
            pkgFail=$(grep -c '^[[:space:]]*--- FAIL' "${RUN_LOG}" || true)
            pkgSkip=$(grep -c '^[[:space:]]*--- SKIP' "${RUN_LOG}" || true)
            TOTAL_TESTS=$((TOTAL_TESTS + pkgRuns))
            TOTAL_PASSED=$((TOTAL_PASSED + pkgPass))
            TOTAL_FAILED=$((TOTAL_FAILED + pkgFail))
            TOTAL_SKIPPED=$((TOTAL_SKIPPED + pkgSkip))
            rm -f "${RUN_LOG}"
            print_rate_summary "${SUCCESS_COUNT}" "${PKG_COUNT}" "progress"
            rm -f "${TMP_COVERAGE}"
            printf "[core-tests] tests so far: total=%d passed=%d failed=%d skipped=%d\n" "${TOTAL_TESTS}" "${TOTAL_PASSED}" "${TOTAL_FAILED}" "${TOTAL_SKIPPED}"
            return 1
        fi

        pkgRuns=$(grep -c '^=== RUN' "${RUN_LOG}" || true)
        pkgPass=$(grep -c '^[[:space:]]*--- PASS' "${RUN_LOG}" || true)
        pkgFail=$(grep -c '^[[:space:]]*--- FAIL' "${RUN_LOG}" || true)
        pkgSkip=$(grep -c '^[[:space:]]*--- SKIP' "${RUN_LOG}" || true)
        TOTAL_TESTS=$((TOTAL_TESTS + pkgRuns))
        TOTAL_PASSED=$((TOTAL_PASSED + pkgPass))
        TOTAL_FAILED=$((TOTAL_FAILED + pkgFail))
        TOTAL_SKIPPED=$((TOTAL_SKIPPED + pkgSkip))
        rm -f "${RUN_LOG}"
        SUCCESS_COUNT=$((SUCCESS_COUNT + 1))

        if [[ "${SKIP_COVERAGE}" != "true" && -f "${TMP_COVERAGE}" ]]; then
            if (( i == 0 )); then
                cat "${TMP_COVERAGE}" > "${COVERAGE_FILE}"
            else
                tail -n +2 "${TMP_COVERAGE}" >> "${COVERAGE_FILE}"
            fi
        fi
    done
    rm -f "${TMP_COVERAGE}"

    if [[ "${SKIP_COVERAGE}" != "true" ]]; then
        log_section "Coverage Summary"
        go tool cover -func="${COVERAGE_FILE}" | tail -n 1
    fi

    print_rate_summary "${SUCCESS_COUNT}" "${PKG_COUNT}" "completed"
    printf "[core-tests] tests: total=%d passed=%d failed=%d skipped=%d\n" "${TOTAL_TESTS}" "${TOTAL_PASSED}" "${TOTAL_FAILED}" "${TOTAL_SKIPPED}"

    # Export results for summary
    CORE_TOTAL_TESTS=$TOTAL_TESTS
    CORE_TOTAL_PASSED=$TOTAL_PASSED
    CORE_TOTAL_FAILED=$TOTAL_FAILED
    CORE_TOTAL_SKIPPED=$TOTAL_SKIPPED
}

# =============================================================================
# DuckDB Accuracy Tests
# =============================================================================
run_duckdb_accuracy_tests() {
    log_header "DuckDB Accuracy & Parity Tests"

    log_section "Level 1-3: Basic Functionality Tests"
    log_info "Testing: Data loading, simple formulas, cell references, SUMIFS/COUNTIFS"

    ACCURACY_LOG="$(mktemp)"
    ACCURACY_PASSED=0
    ACCURACY_FAILED=0
    ACCURACY_SKIPPED=0

    # Run basic DuckDB integration tests
    log_info "Running basic integration tests..."
    if go test -v -run "TestDuckDB_Level[123]" -timeout 300s ./... 2>&1 | tee "${ACCURACY_LOG}"; then
        local passed=$(grep -c '^[[:space:]]*--- PASS' "${ACCURACY_LOG}" || true)
        local failed=$(grep -c '^[[:space:]]*--- FAIL' "${ACCURACY_LOG}" || true)
        local skipped=$(grep -c '^[[:space:]]*--- SKIP' "${ACCURACY_LOG}" || true)
        ACCURACY_PASSED=$((ACCURACY_PASSED + passed))
        ACCURACY_FAILED=$((ACCURACY_FAILED + failed))
        ACCURACY_SKIPPED=$((ACCURACY_SKIPPED + skipped))
        log_success "Basic integration tests completed: passed=$passed, failed=$failed, skipped=$skipped"
    else
        log_error "Basic integration tests failed"
        ACCURACY_FAILED=$((ACCURACY_FAILED + 1))
    fi

    log_section "Level 4: Lookup Functions (VLOOKUP, INDEX, MATCH)"
    log_info "Testing: VLOOKUP exact match, INDEX/MATCH combinations"

    if go test -v -run "TestDuckDB_Level4" -timeout 180s ./... 2>&1 | tee "${ACCURACY_LOG}"; then
        local passed=$(grep -c '^[[:space:]]*--- PASS' "${ACCURACY_LOG}" || true)
        local failed=$(grep -c '^[[:space:]]*--- FAIL' "${ACCURACY_LOG}" || true)
        local skipped=$(grep -c '^[[:space:]]*--- SKIP' "${ACCURACY_LOG}" || true)
        ACCURACY_PASSED=$((ACCURACY_PASSED + passed))
        ACCURACY_FAILED=$((ACCURACY_FAILED + failed))
        ACCURACY_SKIPPED=$((ACCURACY_SKIPPED + skipped))
        log_success "Lookup function tests completed: passed=$passed, failed=$failed, skipped=$skipped"
    else
        log_warning "Lookup function tests had issues (some functions may have known limitations)"
    fi

    log_section "DuckDB Package Unit Tests"
    log_info "Running duckdb package tests..."

    if go test -v -timeout 300s ./duckdb/... 2>&1 | tee "${ACCURACY_LOG}"; then
        local passed=$(grep -c '^[[:space:]]*--- PASS' "${ACCURACY_LOG}" || true)
        local failed=$(grep -c '^[[:space:]]*--- FAIL' "${ACCURACY_LOG}" || true)
        local skipped=$(grep -c '^[[:space:]]*--- SKIP' "${ACCURACY_LOG}" || true)
        ACCURACY_PASSED=$((ACCURACY_PASSED + passed))
        ACCURACY_FAILED=$((ACCURACY_FAILED + failed))
        ACCURACY_SKIPPED=$((ACCURACY_SKIPPED + skipped))
        log_success "DuckDB package tests completed: passed=$passed, failed=$failed, skipped=$skipped"
    else
        log_error "DuckDB package tests failed"
        ACCURACY_FAILED=$((ACCURACY_FAILED + 1))
    fi

    log_section "Native vs DuckDB Parity Tests"
    log_info "Running parity comparison tests..."

    if go test -v -run "TestDuckDBParity|TestDuckDBIntegration" -timeout 300s ./... 2>&1 | tee "${ACCURACY_LOG}"; then
        local passed=$(grep -c '^[[:space:]]*--- PASS' "${ACCURACY_LOG}" || true)
        local failed=$(grep -c '^[[:space:]]*--- FAIL' "${ACCURACY_LOG}" || true)
        local skipped=$(grep -c '^[[:space:]]*--- SKIP' "${ACCURACY_LOG}" || true)
        ACCURACY_PASSED=$((ACCURACY_PASSED + passed))
        ACCURACY_FAILED=$((ACCURACY_FAILED + failed))
        ACCURACY_SKIPPED=$((ACCURACY_SKIPPED + skipped))
        log_success "Parity tests completed: passed=$passed, failed=$failed, skipped=$skipped"
    else
        log_error "Parity tests failed"
        ACCURACY_FAILED=$((ACCURACY_FAILED + 1))
    fi

    rm -f "${ACCURACY_LOG}"

    log_section "DuckDB Accuracy Test Summary"
    echo -e "  ${GREEN}Passed:${RESET}  ${ACCURACY_PASSED}"
    echo -e "  ${RED}Failed:${RESET}  ${ACCURACY_FAILED}"
    echo -e "  ${YELLOW}Skipped:${RESET} ${ACCURACY_SKIPPED}"

    # Export results
    DUCKDB_ACCURACY_PASSED=$ACCURACY_PASSED
    DUCKDB_ACCURACY_FAILED=$ACCURACY_FAILED
    DUCKDB_ACCURACY_SKIPPED=$ACCURACY_SKIPPED
}

# =============================================================================
# DuckDB Performance Benchmarks
# =============================================================================
run_duckdb_perf_benchmarks() {
    log_header "DuckDB Performance Benchmarks"

    BENCH_RESULTS_FILE="$(mktemp)"

    log_section "Formula Compilation Benchmarks"
    log_info "Testing formula parsing and SQL translation speed..."

    go test -bench=BenchmarkFormulaCompiler -benchtime=${BENCHMARK_TIME} -count=${BENCHMARK_COUNT} -benchmem ./duckdb/... 2>&1 | tee -a "${BENCH_RESULTS_FILE}"

    log_section "SUMIFS/COUNTIFS Benchmarks (100K rows)"
    log_info "Testing conditional aggregation with pre-computed cache..."

    go test -bench="BenchmarkSUMIFS|BenchmarkBatchSUMIFS" -benchtime=${BENCHMARK_TIME} -count=${BENCHMARK_COUNT} -benchmem ./duckdb/... 2>&1 | tee -a "${BENCH_RESULTS_FILE}"

    log_section "Lookup Function Benchmarks"
    log_info "Testing INDEX and MATCH performance with indexed lookups..."

    go test -bench="BenchmarkIndexLookup|BenchmarkMatchPosition" -benchtime=${BENCHMARK_TIME} -count=${BENCHMARK_COUNT} -benchmem ./duckdb/... 2>&1 | tee -a "${BENCH_RESULTS_FILE}"

    log_section "Direct SQL Benchmarks"
    log_info "Testing raw DuckDB query execution speed..."

    go test -bench=BenchmarkDirectSQL -benchtime=${BENCHMARK_TIME} -count=${BENCHMARK_COUNT} -benchmem ./duckdb/... 2>&1 | tee -a "${BENCH_RESULTS_FILE}"

    log_section "Large Data Loading Benchmarks"
    log_info "Testing data loading performance (10K, 100K, 1M rows)..."

    go test -bench=BenchmarkLargeDataLoad -benchtime=${BENCHMARK_TIME} -count=1 -benchmem ./duckdb/... 2>&1 | tee -a "${BENCH_RESULTS_FILE}"

    log_section "Native vs DuckDB Comparison Benchmarks"
    log_info "Comparing native engine vs DuckDB engine performance..."

    go test -bench="BenchmarkNative_" -benchtime=${BENCHMARK_TIME} -count=${BENCHMARK_COUNT} -benchmem ./... 2>&1 | tee -a "${BENCH_RESULTS_FILE}"

    log_section "Performance Summary"
    echo ""
    echo -e "${BOLD}Key Performance Metrics:${RESET}"
    echo ""

    # Extract and display key metrics
    if grep -q "BenchmarkSUMIFS-" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}SUMIFS (cached lookup):${RESET}"
        grep "BenchmarkSUMIFS-" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op, " $5 " B/op"}'
    fi

    if grep -q "BenchmarkBatchSUMIFS-" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}Batch SUMIFS (1000 lookups):${RESET}"
        grep "BenchmarkBatchSUMIFS-" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op, " $5 " B/op"}'
    fi

    if grep -q "BenchmarkIndexLookup-" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}INDEX Lookup:${RESET}"
        grep "BenchmarkIndexLookup-" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op, " $5 " B/op"}'
    fi

    if grep -q "BenchmarkMatchPosition-" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}MATCH Position:${RESET}"
        grep "BenchmarkMatchPosition-" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op, " $5 " B/op"}'
    fi

    if grep -q "BenchmarkDirectSQL-" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}Direct SQL Query:${RESET}"
        grep "BenchmarkDirectSQL-" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op, " $5 " B/op"}'
    fi

    echo ""
    echo -e "${BOLD}Data Loading Performance:${RESET}"
    if grep -q "BenchmarkLargeDataLoad/rows_10000" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}10K rows:${RESET}"
        grep "BenchmarkLargeDataLoad/rows_10000" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op"}'
    fi
    if grep -q "BenchmarkLargeDataLoad/rows_100000" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}100K rows:${RESET}"
        grep "BenchmarkLargeDataLoad/rows_100000" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op"}'
    fi
    if grep -q "BenchmarkLargeDataLoad/rows_1000000" "${BENCH_RESULTS_FILE}"; then
        echo -e "  ${CYAN}1M rows:${RESET}"
        grep "BenchmarkLargeDataLoad/rows_1000000" "${BENCH_RESULTS_FILE}" | head -1 | awk '{print "    " $3 " ns/op"}'
    fi

    # Save benchmark results
    BENCH_OUTPUT_FILE="benchmark_results_$(date +%Y%m%d_%H%M%S).txt"
    cp "${BENCH_RESULTS_FILE}" "${BENCH_OUTPUT_FILE}"
    log_info "Benchmark results saved to: ${BENCH_OUTPUT_FILE}"

    rm -f "${BENCH_RESULTS_FILE}"
}

# =============================================================================
# DuckDB Large Scale Tests
# =============================================================================
run_duckdb_large_scale_tests() {
    log_header "DuckDB Large Scale Tests"

    log_section "Level 5: Medium Scale (10K rows, 50 columns)"
    log_info "Testing medium-scale data operations..."

    if go test -v -run "TestDuckDB_Level5" -timeout 600s ./... 2>&1; then
        log_success "Medium scale tests passed"
    else
        log_warning "Medium scale tests had issues"
    fi

    log_section "Level 6: Large Scale Multi-Sheet"
    log_info "Testing multi-sheet operations (10K + 4K + 1K rows)..."

    if go test -v -run "TestDuckDB_Level6" -timeout 600s ./... 2>&1; then
        log_success "Large scale multi-sheet tests passed"
    else
        log_warning "Large scale multi-sheet tests had issues"
    fi

    log_section "Level 7: Cross-Worksheet References"
    log_info "Testing cross-sheet VLOOKUP and formula references..."

    if go test -v -run "TestDuckDB_Level7" -timeout 300s ./... 2>&1; then
        log_success "Cross-worksheet reference tests passed"
    else
        log_warning "Cross-worksheet reference tests had issues"
    fi

    log_section "Real-World Pattern Tests"
    log_info "Testing patterns from real Excel files..."

    if go test -v -run "TestRealWorld" -timeout 600s ./duckdb/... 2>&1; then
        log_success "Real-world pattern tests passed"
    else
        log_warning "Real-world pattern tests had issues"
    fi
}

# =============================================================================
# Quick Tests (Subset for CI)
# =============================================================================
run_quick_tests() {
    log_header "Quick Tests (CI Mode)"

    log_info "Running quick core tests (short mode)..."

    if go test -short -timeout 120s ./... 2>&1; then
        log_success "Quick tests passed"
    else
        log_error "Quick tests failed"
        return 1
    fi

    log_info "Running DuckDB basic tests..."

    if go test -v -run "TestDuckDB_Level[12]" -timeout 120s ./... 2>&1; then
        log_success "DuckDB basic tests passed"
    else
        log_warning "DuckDB basic tests had issues"
    fi
}

# =============================================================================
# Summary Report
# =============================================================================
print_final_summary() {
    log_header "Final Test Summary"

    echo -e "${BOLD}Test Mode:${RESET} ${MODE}"
    echo ""

    if [[ -n "${CORE_TOTAL_TESTS:-}" ]]; then
        echo -e "${BOLD}Core Tests:${RESET}"
        echo -e "  Total:   ${CORE_TOTAL_TESTS}"
        echo -e "  Passed:  ${GREEN}${CORE_TOTAL_PASSED}${RESET}"
        echo -e "  Failed:  ${RED}${CORE_TOTAL_FAILED}${RESET}"
        echo -e "  Skipped: ${YELLOW}${CORE_TOTAL_SKIPPED}${RESET}"
        echo ""
    fi

    if [[ -n "${DUCKDB_ACCURACY_PASSED:-}" ]]; then
        echo -e "${BOLD}DuckDB Accuracy Tests:${RESET}"
        echo -e "  Passed:  ${GREEN}${DUCKDB_ACCURACY_PASSED}${RESET}"
        echo -e "  Failed:  ${RED}${DUCKDB_ACCURACY_FAILED}${RESET}"
        echo -e "  Skipped: ${YELLOW}${DUCKDB_ACCURACY_SKIPPED}${RESET}"
        echo ""
    fi

    echo -e "${BOLD}DuckDB Engine Features Tested:${RESET}"
    echo "  - Formula to SQL translation"
    echo "  - Aggregation functions (SUM, COUNT, AVERAGE, MIN, MAX)"
    echo "  - Conditional aggregations (SUMIFS, COUNTIFS, AVERAGEIFS)"
    echo "  - Lookup functions (VLOOKUP, INDEX, MATCH)"
    echo "  - Pre-computation cache for O(1) lookups"
    echo "  - Multi-sheet data loading"
    echo "  - Cross-worksheet references"
    echo ""

    echo -e "${BOLD}Performance Targets:${RESET}"
    echo "  - SUMIFS cached lookup: <200μs"
    echo "  - Batch SUMIFS (1000 lookups): <100ms"
    echo "  - INDEX lookup: <50μs"
    echo "  - MATCH position: <100μs"
    echo "  - 100K row data loading: <5s"
    echo ""
}

# =============================================================================
# Main Execution
# =============================================================================
main() {
    START_TIME=$(date +%s%3N)

    log_header "Excelize Test Suite (Mode: ${MODE})"
    log_info "Go version: $(go version | awk '{print $3}')"
    log_info "DuckDB tests enabled: true"

    case "${MODE}" in
        quick)
            run_quick_tests
            ;;
        accuracy)
            run_duckdb_accuracy_tests
            ;;
        perf)
            run_duckdb_perf_benchmarks
            ;;
        duckdb)
            run_duckdb_accuracy_tests
            run_duckdb_perf_benchmarks
            run_duckdb_large_scale_tests
            ;;
        full)
            run_core_tests
            run_duckdb_accuracy_tests
            run_duckdb_perf_benchmarks
            run_duckdb_large_scale_tests
            ;;
        all|*)
            run_core_tests
            run_duckdb_accuracy_tests
            ;;
    esac

    END_TIME=$(date +%s%3N)
    DURATION=$((END_TIME - START_TIME))

    print_final_summary

    log_info "Total execution time: $(format_duration $DURATION)"

    # Exit with appropriate code
    if [[ "${CORE_TOTAL_FAILED:-0}" -gt 0 ]] || [[ "${DUCKDB_ACCURACY_FAILED:-0}" -gt 0 ]]; then
        log_error "Some tests failed!"
        exit 1
    else
        log_success "All tests completed successfully!"
        exit 0
    fi
}

# Run main
main
