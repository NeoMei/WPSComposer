tell application "/Applications/Microsoft Excel.app"
if (name of every workbook) is not {"工作簿1"} then error "unexpected inventory"
set ownedBook to open workbook workbook file name "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx" with editable
return {name of ownedBook, full name of ownedBook, name of every workbook}
end tell
