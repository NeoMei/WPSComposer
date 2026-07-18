# Installed-Mac conversion fixtures

The real conversion gate intentionally uses files serialized by WPS rather
than adding synthetic Office binaries to this repository.

For `.doc`, `.docx`, `.xls`, `.ppt`, and `.pptx`, use the matching files under:

```text
/Applications/wpsoffice.app/Contents/Resources/office6/mui/default/templates/oleNewFile/
```

The two-visible-sheet `.xlsx` gate fixture is derived from WPS's bundled
`newfile.xlsx`: it retains the original WPS package parts, adds a second visible
worksheet relationship and content-type entry, and gives both worksheets
distinct inline-string content. Before conversion, confirm `xl/workbook.xml`
contains two `<sheet ... state="visible">` entries. After conversion,
`pdfinfo` must report two pages.

Run all six conversions twice into different output directories. A run passes
only when every PDF has a valid `%PDF-` signature, is at least 1,024 bytes,
registration is restored, no WPS staging session remains, and no save dialog or
new permission interaction appears.
