tell application "/Applications/Microsoft Excel.app"
if (count of workbooks) is not 0 then error "unexpected books before reopen"
set ownedBook to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/isolated-launch.xlsx" with editable
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/isolated-launch.xlsx" then error "wrong reopened path"
if name of every worksheet of ownedBook is not {"Data", "Results"} then error "wrong reopened sheets"
if value of range "Z40" of worksheet 1 of ownedBook is not "ISOLATED-BOUNDARY-CHECK" then error "lost marker"
close ownedBook saving no
if (count of workbooks) is not 0 then error "not closed"
if display alerts is not true then error "alerts not restored"
quit saving no
return "DELETE_SAVE_REOPEN_CLEANUP_PASS"
end tell
