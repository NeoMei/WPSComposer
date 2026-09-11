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

tell application "/Applications/Microsoft Excel.app"
set probeWorkbookRows to {}
repeat with probeIndex from 1 to count workbooks
set probeBook to workbook probeIndex
set end of probeWorkbookRows to {name of probeBook,full name of probeBook,saved of probeBook,value of range "A1" of worksheet 1 of probeBook}
end repeat
return "{" & "\"books\":" & my j(probeWorkbookRows) & ",\"alerts\":" & my j(display alerts) & "}"
end tell
