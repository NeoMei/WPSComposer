tell application "/Applications/Microsoft Excel.app"
if (name of every workbook) is not {"工作簿1"} then error "unexpected inventory"
open POSIX file "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-existing-delete-diagnosis/existing-owned.xlsx"
return name of every workbook
end tell
