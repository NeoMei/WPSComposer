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
set ownedBook to workbook "owned-2840167844f74e6ea6ab3df844ff3728.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-prkz6bw3/owned-2840167844f74e6ea6ab3df844ff3728.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to shape "StableRevenue" of ws
set acceptedKeys to {}
set rejectedKeys to {}
set left position of obj to 35
set end of acceptedKeys to "geometry.left"
set fore color of fill format of obj to {240, 240, 240}
set end of acceptedKeys to "fill.color"
set transparency of fill format of obj to 0.1
set end of acceptedKeys to "fill.transparency"
set fore color of line format of obj to {18, 52, 86}
set end of acceptedKeys to "line.color"
set weight of line format of obj to 2
set end of acceptedKeys to "line.weight"
set visible of line format of obj to true
set end of acceptedKeys to "line.visible"
return "{" & "\"accepted\":" & my j(acceptedKeys) & ",\"rejected\":" & my j(rejectedKeys) & "}"
 end tell
end timeout
