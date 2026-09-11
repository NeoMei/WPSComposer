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
if not application "/Applications/Microsoft Excel.app" is running then return "{\"active\":false}"
with timeout of 60 seconds
 tell application "/Applications/Microsoft Excel.app"

if (count of workbooks) is 0 then return "{\"active\":false}"
set candidate to active workbook
return "{" & "\"active\":" & my j(true) & ",\"path\":" & my j(full name of candidate) & ",\"name\":" & my j(name of candidate) & ",\"read_only\":" & my j(read only of candidate) & "}"
 end tell
end timeout
