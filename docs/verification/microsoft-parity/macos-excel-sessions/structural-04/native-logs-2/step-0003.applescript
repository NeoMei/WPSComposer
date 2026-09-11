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
set ownedBook to workbook "owned-7259121f846549fd9e5b781983bfba02.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-vp3wuouo/owned-7259121f846549fd9e5b781983bfba02.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to range "B2:D4" of ws
set obj to range "B2" of ws
cut range (entire row of obj)
insert into range (range "7:7" of ws)
set cut copy mode to false
return "{" & "\"type\":" & my j("row") & ",\"moved\":" & my j(true) & ",\"from\":" & my j("sheet:1/range:B2:D4") & ",\"row\":" & my j(6) & ",\"clipboard_changed\":" & my j(true) & ",\"requested_index\":" & my j(7) & "}"
 end tell
end timeout
