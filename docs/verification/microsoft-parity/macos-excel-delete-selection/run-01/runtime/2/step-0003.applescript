use framework "Foundation"
use scripting additions
on j(v)
 if v is missing value then return "null"
 try
  set box to current application's NSArray's arrayWithObject:v
  set d to current application's NSJSONSerialization's dataWithJSONObject:box options:0 |error|:(missing value)
  set s to (current application's NSString's alloc()'s initWithData:d encoding:4) as text
  return text 2 thru -2 of s
 on error
  try
   return my j(v as text)
  on error
   return "null"
  end try
 end try
end j
on joined(itemsList)
 set oldDelimiters to AppleScript's text item delimiters
 set AppleScript's text item delimiters to ","
 set resultText to itemsList as text
 set AppleScript's text item delimiters to oldDelimiters
 return resultText
end joined
with timeout of 60 seconds
 tell application "/Applications/Microsoft Excel.app"
set ownedBook to workbook "owned-26996a7da65040c48f6d19ecdacbfc0d.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-7vp9qy8_/owned-26996a7da65040c48f6d19ecdacbfc0d.xlsx" then error "Owned workbook identity mismatch"
if "DeleteProbeA" is in (name of every worksheet of ownedBook) then error "Duplicate sheet name"
tell ownedBook
 make new worksheet at end with properties {name:"DeleteProbeA"}
end tell
return "{" & "\"type\":" & my j("sheet") & ",\"path\":" & my j("sheet:" & (count of worksheets of ownedBook)) & "}"
 end tell
end timeout
