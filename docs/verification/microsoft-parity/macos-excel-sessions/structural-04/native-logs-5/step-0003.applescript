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
set ownedBook to workbook "owned-1c528e301be04bcdb81d25d8806f3fb2.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-e8j6qyvy/owned-1c528e301be04bcdb81d25d8806f3fb2.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to range "B2:D4" of ws
set obj to range "B2" of ws
cut range (entire column of obj)
insert into range (range "E:E" of ws)
set cut copy mode to false
return "{" & "\"type\":" & my j("column") & ",\"moved\":" & my j(true) & ",\"from\":" & my j("sheet:1/range:B2:D4") & ",\"column\":" & my j(4) & ",\"clipboard_changed\":" & my j(true) & ",\"requested_index\":" & my j(5) & "}"
 end tell
end timeout
