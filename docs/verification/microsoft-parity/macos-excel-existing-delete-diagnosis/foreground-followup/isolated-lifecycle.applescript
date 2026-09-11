tell application "/Applications/Microsoft Excel.app"
if name of every workbook is not {"isolated-launch.xlsx"} then error "wrong inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/excel-command-boundary-20260910/isolated-launch.xlsx" then error "wrong path"
set value of range "Z40" of worksheet 1 of ownedBook to "ISOLATED-BOUNDARY-CHECK"
save ownedBook
return {saved of ownedBook, value of range "Z40" of worksheet 1 of ownedBook}
end tell
