with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set b to workbook "active-owned.xlsx"
if full name of b is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-active-fixture-nzxj9bpg/active-owned.xlsx" then error "identity"
return value of range "B2" of worksheet "Data" of b
end tell
end timeout