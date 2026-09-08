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
set ownedBook to workbook "owned-70eda9be2f18451181f170ed2397e128.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-82dy8x6o/owned-70eda9be2f18451181f170ed2397e128.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to range "A1:B4" of ws
set acceptedKeys to {}
set rejectedKeys to {}
set ownedBorder to get border obj which border edge left
set line style of ownedBorder to continuous
set weight of ownedBorder to border weight thin
set color of ownedBorder to {18, 52, 86}
set end of acceptedKeys to "borders.7"
set ownedBorder to get border obj which border edge top
set line style of ownedBorder to continuous
set weight of ownedBorder to border weight thin
set color of ownedBorder to {18, 52, 86}
set end of acceptedKeys to "borders.8"
set ownedBorder to get border obj which border edge bottom
set line style of ownedBorder to continuous
set weight of ownedBorder to border weight thin
set color of ownedBorder to {18, 52, 86}
set end of acceptedKeys to "borders.9"
set ownedBorder to get border obj which border edge right
set line style of ownedBorder to continuous
set weight of ownedBorder to border weight thin
set color of ownedBorder to {18, 52, 86}
set end of acceptedKeys to "borders.10"
set ownedBorder to get border obj which border inside vertical
set line style of ownedBorder to continuous
set weight of ownedBorder to border weight thin
set color of ownedBorder to {18, 52, 86}
set end of acceptedKeys to "borders.11"
set ownedBorder to get border obj which border inside horizontal
set line style of ownedBorder to continuous
set weight of ownedBorder to border weight thin
set color of ownedBorder to {18, 52, 86}
set end of acceptedKeys to "borders.12"
calculate obj
return "{" & "\"accepted\":" & my j(acceptedKeys) & ",\"rejected\":" & my j(rejectedKeys) & "}"
 end tell
end timeout
