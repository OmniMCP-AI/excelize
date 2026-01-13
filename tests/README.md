# tests

This directory captures high level validation assets that do not belong in Go unit / integration tests.

## Layout
- `manual/` - curated manual test plans and sample workbooks that mirror how QA staff execute Excel compatibility tests.
  - `excel2016_system_test_plan.md` - test plan imported from the user provided Excel 2016 document. The plan keeps the original Chinese wording so QA can continue to work with an authoritative reference.
  - `data/` - folder that stores all binary workbooks referenced by the plan; see `tests/manual/data/README.md` for the manifest.
  - `execution_log_template.csv` - PASS/FAIL tracker template based on the plan's Section 7.4 table; duplicate per run and fill in the results.
  - `execution_log_2024q1.csv` - first populated run log (extend or add files per release).
  - `tools/generate_data.go` - Go generator that rebuilds every workbook in the `data/` folder.

## How to use
1. Open one of the plans (for now `manual/excel2016_system_test_plan.md`).
2. Ensure every workbook documented in `manual/data/README.md` exists before testing so formulas reference identical fixtures (regenerate via `go run tests/manual/tools/generate_data.go` when needed).
3. Copy `manual/execution_log_template.csv` to a dated file (example already provided as `execution_log_2024q1.csv`) and log PASS/FAIL for each case while executing Section 7.
4. When new cases appear, extend the Markdown document or attach new documents under `manual/` so that the historical context lives with the repository instead of a local download.

## Future work
- Track PASS/FAIL status per release by adding CSV/JSON exports next to the plan.
- Attach macros or Go based automation scripts under `tests/automation/` (not yet created) once portions of the plan are automated.
