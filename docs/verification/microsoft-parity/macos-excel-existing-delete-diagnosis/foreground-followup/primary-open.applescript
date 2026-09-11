tell application "/Applications/Microsoft Excel.app"
if name of every workbook is not {"工作簿5", "工作簿7"} then error "Unexpected preexisting workbooks"
set ownedBook to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/owned-command-test.xlsx" with editable
return {name of ownedBook, full name of ownedBook, name of every workbook}
end tell
