tell application "/Applications/Microsoft Excel.app"
if name of every workbook is not {"isolated-launch.xlsx"} then error "wrong inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/isolated-launch.xlsx" then error "wrong path"
if name of every worksheet of ownedBook is not {"Data", "Results", "Business report"} then error "wrong sheets"
if display alerts is not false then error "alerts not suppressed"
try
 delete worksheet "Business report" of ownedBook
 set display alerts to true
on error msg number n
 set display alerts to true
 error msg number n
end try
if name of every worksheet of ownedBook is not {"Data", "Results"} then error "delete not acknowledged"
save ownedBook
if saved of ownedBook is not true then error "not saved"
close ownedBook saving no
if count workbooks is not 0 then error "not closed"
set ownedBook to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/isolated-launch.xlsx" with editable
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/isolated-launch.xlsx" then error "wrong reopened path"
if name of every worksheet of ownedBook is not {"Data", "Results"} then error "wrong reopened sheets"
if value of range "Z40" of worksheet 1 of ownedBook is not "ISOLATED-BOUNDARY-CHECK" then error "lost marker"
close ownedBook saving no
if count workbooks is not 0 then error "not closed"
return {"DELETE_SAVE_REOPEN_PASS", display alerts}
end tell
