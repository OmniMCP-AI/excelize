#!/usr/bin/env bash
set -euo pipefail

COVERAGE_FILE="${COVERAGE_FILE:-coverage.out}"
PKG_PATTERN="${PKG_PATTERN:-./...}"
EXCLUDE_PACKAGES="${EXCLUDE_PACKAGES:-github.com/xuri/excelize/v2/tests/manual/tools}"
GO_TEST_FLAGS="${GO_TEST_FLAGS:--v}"
export GOCACHE="${GOCACHE:-/tmp/go-build}"
GREEN="\033[32m"
YELLOW="\033[33m"
RED="\033[31m"
RESET="\033[0m"

echo "[core-tests] collecting package list (${PKG_PATTERN})"
PKGS=()
EXCLUDES=()
while IFS= read -r line; do
  [[ -z "${line}" ]] && continue
  EXCLUDES+=("${line}")
done < <(printf "%s\n" "${EXCLUDE_PACKAGES}" | tr ' ' '\n' | tr ',' '\n')
PKG_LIST=""
if ! PKG_LIST=$(go list "${PKG_PATTERN}"); then
  echo "[core-tests] failed to list packages for pattern ${PKG_PATTERN}"
  exit 1
fi
while IFS= read -r pkg; do
  skip=false
  for exclude in "${EXCLUDES[@]:-}"; do
    if [[ "${pkg}" == "${exclude}" ]]; then
      echo "[core-tests] skipping excluded package ${pkg}"
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
  echo "[core-tests] no packages matched pattern ${PKG_PATTERN}"
  exit 1
fi

echo "[core-tests] running ${PKG_COUNT} packages with coverage -> ${COVERAGE_FILE}"
TMP_COVERAGE="$(mktemp)"
: > "${COVERAGE_FILE}"
SUCCESS_COUNT=0
TOTAL_TESTS=0
TOTAL_PASSED=0
TOTAL_FAILED=0
TOTAL_SKIPPED=0

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

for ((i=0; i<PKG_COUNT; i++)); do
  pkg="${PKGS[$i]}"
  printf '[core-tests] (%d/%d) go test %s\n' "$((i+1))" "$PKG_COUNT" "$pkg"
  RUN_LOG="$(mktemp)"
  if ! GOFLAGS="${GOFLAGS:-}" go test $GO_TEST_FLAGS "$pkg" -covermode=atomic -coverprofile="${TMP_COVERAGE}" 2>&1 | tee "${RUN_LOG}"; then
    printf "%b[core-tests] FAILED in package %s%b\n" "${RED}" "${pkg}" "${RESET}"
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
    exit 1
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
  if (( i == 0 )); then
    cat "${TMP_COVERAGE}" > "${COVERAGE_FILE}"
  else
    tail -n +2 "${TMP_COVERAGE}" >> "${COVERAGE_FILE}"
  fi
done
rm -f "${TMP_COVERAGE}"

echo "[core-tests] coverage summary"
go tool cover -func="${COVERAGE_FILE}" | tail -n 1
print_rate_summary "${SUCCESS_COUNT}" "${PKG_COUNT}" "completed"
printf "[core-tests] tests: total=%d passed=%d failed=%d skipped=%d\n" "${TOTAL_TESTS}" "${TOTAL_PASSED}" "${TOTAL_FAILED}" "${TOTAL_SKIPPED}"
