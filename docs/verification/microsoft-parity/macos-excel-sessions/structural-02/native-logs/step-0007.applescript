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
set ownedBook to workbook "owned-ccdf820e0cc041ce84f6aac0b7e684aa.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-r6ze4jgk/owned-ccdf820e0cc041ce84f6aac0b7e684aa.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to ws
set sourceSheetName to name of obj
set beforeSheetNames to name of every worksheet of ownedBook
set beforeSheetCount to count of worksheets of ownedBook
copy worksheet obj after worksheet 2 of ownedBook
if (count of worksheets of ownedBook) is not beforeSheetCount + 1 then error "Worksheet clone count mismatch"
set resultingSheetName to name of worksheet (2 + 1) of ownedBook
set resultingIndex to entry_index of worksheet resultingSheetName of ownedBook
return "{" & "\"type\":" & my j("sheet") & ",\"cloned\":" & my j(true) & ",\"from\":" & my j("sheet:1") & ",\"path\":" & my j("sheet:" & resultingIndex) & "}"
 end tell
end timeout
