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
set ownedBook to workbook "owned-e3e71de7f72d4182a807d43b9b1b7d54.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-j3igs8nx/owned-e3e71de7f72d4182a807d43b9b1b7d54.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to ws
set obj to range "A1" of ws
set value of obj to "C1"
set font size of font object of obj to 11
set bold of font object of obj to true
set color of interior object of obj to {68, 114, 196}
set color of font object of obj to {255, 255, 255}
set obj to range "B1" of ws
set value of obj to "C2"
set font size of font object of obj to 11
set bold of font object of obj to true
set color of interior object of obj to {68, 114, 196}
set color of font object of obj to {255, 255, 255}
set obj to range "C1" of ws
set value of obj to "C3"
set font size of font object of obj to 11
set bold of font object of obj to true
set color of interior object of obj to {68, 114, 196}
set color of font object of obj to {255, 255, 255}
set obj to range "D1" of ws
set value of obj to "C4"
set font size of font object of obj to 11
set bold of font object of obj to true
set color of interior object of obj to {68, 114, 196}
set color of font object of obj to {255, 255, 255}
set obj to range "E1" of ws
set value of obj to "C5"
set font size of font object of obj to 11
set bold of font object of obj to true
set color of interior object of obj to {68, 114, 196}
set color of font object of obj to {255, 255, 255}
set obj to range "A2" of ws
set value of obj to "R2"
set font size of font object of obj to 11
set obj to range "B2" of ws
set value of obj to 22
set font size of font object of obj to 11
set obj to range "C2" of ws
set value of obj to 23
set font size of font object of obj to 11
set obj to range "D2" of ws
set value of obj to 24
set font size of font object of obj to 11
set obj to range "E2" of ws
set value of obj to 25
set font size of font object of obj to 11
set obj to range "A3" of ws
set value of obj to "R3"
set font size of font object of obj to 11
set obj to range "B3" of ws
set value of obj to 32
set font size of font object of obj to 11
set obj to range "C3" of ws
set value of obj to 33
set font size of font object of obj to 11
set obj to range "D3" of ws
set value of obj to 34
set font size of font object of obj to 11
set obj to range "E3" of ws
set value of obj to 35
set font size of font object of obj to 11
set obj to range "A4" of ws
set value of obj to "R4"
set font size of font object of obj to 11
set obj to range "B4" of ws
set value of obj to 42
set font size of font object of obj to 11
set obj to range "C4" of ws
set value of obj to 43
set font size of font object of obj to 11
set obj to range "D4" of ws
set value of obj to 44
set font size of font object of obj to 11
set obj to range "E4" of ws
set value of obj to 45
set font size of font object of obj to 11
set obj to range "A5" of ws
set value of obj to "R5"
set font size of font object of obj to 11
set obj to range "B5" of ws
set value of obj to 52
set font size of font object of obj to 11
set obj to range "C5" of ws
set value of obj to 53
set font size of font object of obj to 11
set obj to range "D5" of ws
set value of obj to 54
set font size of font object of obj to 11
set obj to range "E5" of ws
set value of obj to 55
set font size of font object of obj to 11
set obj to range "A6" of ws
set value of obj to "R6"
set font size of font object of obj to 11
set obj to range "B6" of ws
set value of obj to 62
set font size of font object of obj to 11
set obj to range "C6" of ws
set value of obj to 63
set font size of font object of obj to 11
set obj to range "D6" of ws
set value of obj to 64
set font size of font object of obj to 11
set obj to range "E6" of ws
set value of obj to 65
set font size of font object of obj to 11
set obj to range "A7" of ws
set value of obj to "R7"
set font size of font object of obj to 11
set obj to range "B7" of ws
set value of obj to 72
set font size of font object of obj to 11
set obj to range "C7" of ws
set value of obj to 73
set font size of font object of obj to 11
set obj to range "D7" of ws
set value of obj to 74
set font size of font object of obj to 11
set obj to range "E7" of ws
set value of obj to 75
set font size of font object of obj to 11
set obj to range "A8" of ws
set value of obj to "R8"
set font size of font object of obj to 11
set obj to range "B8" of ws
set value of obj to 82
set font size of font object of obj to 11
set obj to range "C8" of ws
set value of obj to 83
set font size of font object of obj to 11
set obj to range "D8" of ws
set value of obj to 84
set font size of font object of obj to 11
set obj to range "E8" of ws
set value of obj to 85
set font size of font object of obj to 11
calculate ws
return "{}"
 end tell
end timeout
