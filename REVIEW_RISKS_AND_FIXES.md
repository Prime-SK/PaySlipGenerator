# PaySlip Parsing Review - Risks and Fixes

## Scope

Reviewed the employee extraction and value parsing logic in n_payslip.py, with focus on the employee number conversion change and downstream numeric handling.

Tested with:

- 111.xlsx
- 222.xlsx

## Summary

Your replacement for employee number parsing is directionally good because it accepts string-based IDs that are still numeric.

However, the direct int(str(no).strip()) approach can reject valid Excel numeric values like 1.0. The code has now been hardened with safe parsers.

## What can go wrong

### 1) Employee rows can be skipped when No is 1.0 or formatted text

How it fails:

- Excel may store No as float values such as 1.0.
- int(str(1.0)) becomes int("1.0"), which raises ValueError.
- Valid employee rows get skipped.

Fix applied:

- Added parse_int_like(value).
- It accepts int, integer-like float, and integer-like strings (including values with commas and spaces).
- It safely rejects non-integer content.

### 2) Gang leader allowance check can crash on non-numeric text

How it fails:

- Existing check used float(emp['gang_leader_allowance'] or 0).
- If this cell contains text or formatted string not directly float-compatible, it can raise ValueError and stop PDF generation.

Fix applied:

- Added parse_float_like(value, default=...).
- Replaced direct float conversion with parse_float_like(..., default=0.0).
- Non-numeric values now safely behave as 0.

### 3) Number formatting can be inconsistent for comma-formatted numeric strings

How it fails:

- Values like 1,234.50 may not parse with direct float(value) in all input forms.
- Output can show raw strings instead of consistent currency formatting.

Fix applied:

- fmt() now uses parse_float_like.
- Numeric-like strings are normalized and formatted consistently.

## Additional risks still present (not changed)

### 4) Header detection can break if column wording changes too much

Details:

- Header mapping depends on keyword fragments in COLUMN_MAP.
- If source sheets rename columns significantly, mapping may fail.

Mitigation:

- Keep COLUMN_MAP updated with new synonym keywords.
- Optionally add a fallback map file or per-client profile.

### 5) Date range extraction assumes both dates are in one text segment

Details:

- Current regex expects a range pattern in one cell string.
- If dates are split across separate cells, extraction may return N/A.

Mitigation:

- Extend extraction logic to combine nearby cells on a row before regex matching.

### 6) Data quality issues in salary fields may still produce unexpected display text

Details:

- Non-numeric salary cells are currently displayed as-is by fmt when parsing fails.

Mitigation:

- Add strict validation mode that reports and skips bad rows with a warning log.

## Verification results after fixes

- 111.xlsx: 34 employees generated successfully.
- 222.xlsx: 48 employees generated successfully.

No regressions observed with your two sample files.

## Recommendation

The current parser is now safer for mixed Excel typing in No and allowance fields. If you expect more workbook variants from multiple sources, the next high-value upgrade is adding a validation report (bad row index, bad field, reason) before PDF generation.
