tell application "/Applications/Microsoft Excel.app"
if name of every workbook is not {"工作簿1"} then error "unexpected initial workbook"
set ownedBook to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/foreground-open.xlsx" with editable
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/foreground-open.xlsx" then error "open mismatch"
set value of range "Z40" of worksheet 1 of ownedBook to "FOREGROUND-OPEN-CHECK"
save ownedBook
if saved of ownedBook is not true then error "save mismatch"
close ownedBook saving no
if name of every workbook is not {"工作簿1"} then error "unexpected remaining workbook"
if saved of workbook 1 is not true then error "blank workbook is dirty"
close workbook 1 saving no
if (count of workbooks) is not 0 then error "not empty"
quit saving no
return "FOREGROUND_OPEN_SAVE_CLOSE_PASS"
end tell
