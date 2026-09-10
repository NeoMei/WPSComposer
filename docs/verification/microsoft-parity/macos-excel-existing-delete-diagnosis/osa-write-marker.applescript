tell application "/Applications/Microsoft Excel.app"
if (name of every workbook) is not {"existing-owned.xlsx"} then error "unexpected inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx" then error "path mismatch"
set probeCell to range "Z40" of worksheet 1 of ownedBook
set beforeValue to value of probeCell
set value of probeCell to "PID-WRITE-PROBE-99513"
set afterValue to value of probeCell
return {beforeValue, afterValue, saved of ownedBook}
end tell
