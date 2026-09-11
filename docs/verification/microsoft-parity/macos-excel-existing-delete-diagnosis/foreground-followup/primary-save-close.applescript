tell application "/Applications/Microsoft Excel.app"
set ownedBook to workbook "owned-command-test.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/owned-command-test.xlsx" then error "wrong owned path"
set value of range "Z40" of worksheet 1 of ownedBook to "PRIMARY-SAVE-CHECK"
save ownedBook
if saved of ownedBook is not true then error "save not acknowledged"
close ownedBook saving no
if name of every workbook is not {"工作簿5", "工作簿7"} then error "cleanup mismatch"
return {display alerts, name of every workbook}
end tell
