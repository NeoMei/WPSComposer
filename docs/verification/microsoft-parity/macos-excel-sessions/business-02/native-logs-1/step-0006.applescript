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
set ownedBook to workbook "owned-2f80ba90bdd84884a149c1f44a036ecb.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-us27yjva/owned-2f80ba90bdd84884a149c1f44a036ecb.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to range "B2" of ws
set acceptedKeys to {}
set rejectedKeys to {}
set bold of font object of obj to true
set end of acceptedKeys to "font.bold"
set italic of font object of obj to true
set end of acceptedKeys to "font.italic"
set font size of font object of obj to 13
set end of acceptedKeys to "font.size"
set color of font object of obj to {18, 52, 86}
set end of acceptedKeys to "font.color"
set color of interior object of obj to {255, 242, 204}
set end of acceptedKeys to "fill.color"
set horizontal alignment of obj to horizontal align center
set end of acceptedKeys to "horizontal_alignment"
set number format of obj to "0.00"
set end of acceptedKeys to "number_format"
calculate obj
return "{" & "\"accepted\":" & my j(acceptedKeys) & ",\"rejected\":" & my j(rejectedKeys) & "}"
 end tell
end timeout
