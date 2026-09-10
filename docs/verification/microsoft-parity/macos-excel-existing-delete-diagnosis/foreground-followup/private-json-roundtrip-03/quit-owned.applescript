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
if (count of workbooks) is not 0 then error "not empty"
if display alerts is not true then error "alerts changed"
quit saving no
return "{\"quit\":true}"
end tell
