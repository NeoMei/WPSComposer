with timeout of 30 seconds
tell application "/Applications/Microsoft Excel.app"
set b to workbook "owned-5f0c410c2c4342f4940f63961bf6ff79.xlsx"
if full name of b is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-ktpexju1/owned-5f0c410c2c4342f4940f63961bf6ff79.xlsx" then error "Owned identity mismatch"
save b
close b saving no
set resultText to ""
set beforeNames to name of every workbook
repeat with bookName in beforeNames
set b to workbook (bookName as text)
set documentPath to full name of b
if documentPath does not start with "/" then
if (name of b) is not "工作簿5" and (name of b) is not "工作簿7" then error "Unrecognized unsaved workbook"
activate object worksheet 1 of b
set token to value of range "A1" of worksheet 1 of b
if ((name of b) is "工作簿5" and token is not "sentinel-58d9ceb7df8e") or ((name of b) is "工作簿7" and token is not "sentinel-e9a7d3dce1ad") then error "Sentinel identity mismatch"
if saved of b then error "Sentinel saved state mismatch"
set resultText to resultText & (name of b) & "|" & token & "|false" & linefeed
end if
if documentPath starts with "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-ktpexju1/" then error "Owned document still open"
end repeat
return resultText
end tell
end timeout