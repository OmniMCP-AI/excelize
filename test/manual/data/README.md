# Manual test data sets

The Excel 2016 system test plan (see `tests/manual/excel2016_system_test_plan.md`) references a few curated workbooks. Create them once and reuse them across runs so manual execution is consistent.

## Required workbooks

| File | Sheets | Notes |
|------|--------|-------|
| `test_data_numeric.xlsx` | Sheet `Numbers` | Populate with the values listed in Section 6.1 (正/负整数, 小数, 零值, 极值). |
| `test_data_text.xlsx` | Sheet `Text` | Include Chinese/English labels, numeric strings, whitespace, newline samples, and special characters (*, ?, ~, &). |
| `test_data_date.xlsx` | Sheet `Dates` | Add coverage for standard dates, leap years, boundary values, and time-of-day entries. |
| `test_data_mixed.xlsx` | Sheet `Mixed` | Compose by blocks of 10 rows: numbers, text, dates, blanks, and error values (#DIV/0!, #N/A, #VALUE!). |
| `test_data_large.xlsx` | Sheets `Values`, `Text`, `Lookup`, `Formulas` | Generate the large synthetic data volumes from Section 6.2. Use scripts if possible so the workbook can be rebuilt deterministically. |
| `test_data_business.xlsx` | Sheets `Sales`, `Inventory`, `Finance`, `Attendance`, `Projects` | Reflect the business scenarios from Section 6.3.

> Store binary files (xlsx/xlsm) in this folder and keep them under version control so the QA team shares a single truth source.

## Naming rules
- Prefer lowercase with underscores as shown above.
- Version files by suffixing `_v2` only when the schema changes; otherwise amend in-place.

## Rebuilding guidance
1. Run `go run tests/manual/tools/generate_data.go` to recreate every workbook deterministically.
2. If manual tweaks are required, fill in the tables exactly as defined in the system test plan to avoid mismatched expectations.
3. When artificial data is required (e.g. 100k rows), keep the generator script under `tests/manual/tools/` (already provided) and note any custom steps in the workbook.

## Checklist
- [ ] All six workbooks exist.
- [ ] Each sheet tab name matches the table above.
- [ ] Macros/scripts included where required (e.g. automation helpers).
- [ ] File metadata documents the generation date and maintainer (Workbook → Info → Properties).

Update this README whenever a new dataset is added so contributors know which artifacts should be present.
