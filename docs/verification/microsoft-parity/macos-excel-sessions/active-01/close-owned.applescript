with timeout of 60 seconds
tell application "/Applications/Microsoft Excel.app"
set b to workbook "active-owned.xlsx"
if full name of b is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-active-fixture-nzxj9bpg/active-owned.xlsx" then error "identity"
close b saving no
return "closed owned"
end tell
end timeout