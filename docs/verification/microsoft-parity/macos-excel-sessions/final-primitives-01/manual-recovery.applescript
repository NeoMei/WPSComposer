with timeout of 30 seconds
tell application "/Applications/Microsoft Excel.app"
set b to workbook "owned-e145a09d9a974b9eaf351722bd3ffb09.xlsx"
if full name of b is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-tnbqumki/owned-e145a09d9a974b9eaf351722bd3ffb09.xlsx" then error "Owned identity mismatch"
save b
close b saving no
set resultText to ""
repeat with bookName in (name of every workbook)
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
if documentPath starts with "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-tnbqumki/" then error "Owned document still open"
end repeat
return resultText
end tell
end timeout