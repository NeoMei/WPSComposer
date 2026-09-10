tell application "/Applications/Microsoft Excel.app"
if (name of every workbook) is not {"existing-owned.xlsx"} then error "unexpected inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx" then error "path mismatch"
if value of range "Z40" of worksheet 1 of ownedBook is not "PID-WRITE-PROBE-99513" then error "marker mismatch"
close ownedBook saving no
if (count of workbooks) is not 0 then error "close did not remove owned workbook"
if display alerts is not true then error "unexpected alerts state"
quit saving no
return "OWNED_CLEANUP_PASS"
end tell
