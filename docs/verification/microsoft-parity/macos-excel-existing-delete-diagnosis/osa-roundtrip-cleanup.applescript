tell application "/Applications/Microsoft Excel.app"
if (name of every workbook) is not {"existing-owned.xlsx"} then error "unexpected inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx" then error "path mismatch"
if value of range "Z40" of worksheet 1 of ownedBook is not "PID-WRITE-PROBE-99513" then error "marker mismatch"
if display alerts is not true then error "alerts unexpectedly changed"
save ownedBook
if saved of ownedBook is not true then error "save did not persist"
close ownedBook saving no
if (count of workbooks) is not 0 then error "close did not remove owned workbook"
open POSIX file "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx"
if (name of every workbook) is not {"existing-owned.xlsx"} then error "unexpected reopen inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx" then error "reopen path mismatch"
if value of range "Z40" of worksheet 1 of ownedBook is not "PID-WRITE-PROBE-99513" then error "reopen marker mismatch"
if (name of every worksheet of ownedBook) is not {"Data", "Results", "Business report"} then error "reopen sheets mismatch"
close ownedBook saving no
if (count of workbooks) is not 0 then error "cannot quit with remaining workbooks"
if display alerts is not true then error "alerts not restored"
quit saving no
return "OWNED_ROUNDTRIP_AND_CLEANUP_PASS"
end tell
