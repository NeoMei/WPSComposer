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
if not application "/Applications/Microsoft Excel.app" is running then error "Bound Excel application is no longer running"
with timeout of 60 seconds
 tell application "/Applications/Microsoft Excel.app"
set ownedBook to workbook "active-owned.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-active-fixture-nzxj9bpg/active-owned.xlsx" then error "Owned workbook identity mismatch"
if full name of active workbook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-active-fixture-nzxj9bpg/active-owned.xlsx" then error "Selection owner mismatch"
set obj to selection
set acceptedKeys to {}
set rejectedKeys to {}
set value of obj to 17
set end of acceptedKeys to "value"
calculate obj
return "{" & "\"accepted\":" & my j(acceptedKeys) & ",\"rejected\":" & my j(rejectedKeys) & "}"
 end tell
end timeout
