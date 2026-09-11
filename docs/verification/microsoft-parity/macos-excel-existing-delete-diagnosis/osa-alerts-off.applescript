tell application "/Applications/Microsoft Excel.app"
if (name of every workbook) is not {"existing-owned.xlsx"} then error "unexpected inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx" then error "path mismatch"
if (name of every worksheet of ownedBook) is not {"Data", "Results", "Business report"} then error "unexpected sheets"
set display alerts to false
return {display alerts, name of every worksheet of ownedBook}
end tell
