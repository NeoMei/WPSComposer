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
set ownedBook to workbook "owned-48e514c56b754f01998aec0041e08361.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-y645z5k7/owned-48e514c56b754f01998aec0041e08361.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to ws
set obj to range "A1" of ws
set value of obj to "Metric"
set font size of font object of obj to 11
set bold of font object of obj to true
set color of interior object of obj to {68, 114, 196}
set color of font object of obj to {255, 255, 255}
set obj to range "B1" of ws
set value of obj to "Value"
set font size of font object of obj to 11
set bold of font object of obj to true
set color of interior object of obj to {68, 114, 196}
set color of font object of obj to {255, 255, 255}
set obj to range "A2" of ws
set value of obj to "Alpha"
set font size of font object of obj to 11
set obj to range "B2" of ws
set value of obj to 10
set font size of font object of obj to 11
set obj to range "A3" of ws
set value of obj to "Beta"
set font size of font object of obj to 11
set obj to range "B3" of ws
set value of obj to 20
set font size of font object of obj to 11
set obj to range "A4" of ws
set value of obj to "Total"
set font size of font object of obj to 11
set obj to range "B4" of ws
set value of obj to "=SUM(B2:B3)"
set font size of font object of obj to 11
calculate ws
return "{}"
 end tell
end timeout
