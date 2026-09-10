# Caption native run-01 offline validator correction

Date: 2026-09-11 (Asia/Shanghai)

The immutable `native-run-01` report recorded `status: PASS` with nine true checks. Its actual `native_rows`, source hashes, DOCX, PDF, reopen evidence, sentinel cleanup, and inventory evidence are preserved unchanged.

Independent review found that the run-01 `validate_readback` implementation was permissive. It accepted a body with a corrupted prefix or appended text, ignored an extra foreign field, and accepted negative or inconsistent bounds. These were validator false positives; no such corruption was found in the retained run-01 rows.

The repaired validator now requires:

- exactly nine closed, ordered rows: one body, three fields, two bookmarks, two caption paragraphs, and one counts row;
- a fully typed row schema with booleans distinct from integers;
- the exact native body splice observed for the controlled fixture;
- the exact controlled field set and order, with only a localized Heading 1 name variable;
- the exact bookmark names, visible values, nonnegative ordered bounds, containment within the matching paragraph, and ordered adjacent caption paragraph bounds;
- exact caption paragraph text, centering, KeepTogether, explicit KeepWithNext values, and zero table/shape counts.

Offline replay against the immutable run-01 report:

```json
{
  "actual_native_rows_pass_hardened_validator": true,
  "body_corruption_rejected": true,
  "foreign_field_rejected": true,
  "negative_bookmark_rejected": true
}
```

Immutable run-01 hashes:

- `report.json`: `21bf10530fbb02e307b5dad4b796b8f245cca7487d47739dfc2e2aff24b0bf8f`
- `caption.docx`: `699604034e25df45548d9a0fc89e22c6dca594abd6619c1dff7252e6a9221134`
- `caption.pdf`: `86674c950a4a50baddacec6df4282974638afbb5cee18bfd6e1d332ff4881662`

Run-01 remains useful native evidence for its retained actual rows and artifacts, but it is not sufficient as the final fixture acceptance because it did not exercise retained nondefault paragraph/font formatting and exact returned handle/trailing-paragraph bounds. The revised fixture adds those checks for native run-02.
