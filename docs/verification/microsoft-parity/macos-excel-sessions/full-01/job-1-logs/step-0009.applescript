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
set ownedBook to workbook "owned-87fa307d633142aabcbf84bf86b6b695.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-ua3u3wfv/owned-87fa307d633142aabcbf84bf86b6b695.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
insert into range (range "E:E" of ws)
set value of range "E1" of ws to "Inserted column"
set value of range "E2" of ws to 88
return "{" & "\"type\":" & my j("column") & ",\"path\":" & my j("sheet:1") & ",\"column\":" & my j(5) & "}"
 end tell
end timeout
