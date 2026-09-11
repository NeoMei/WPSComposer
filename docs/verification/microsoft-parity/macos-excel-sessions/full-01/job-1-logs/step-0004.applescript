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
activate object ws
set obj to chart object 1 of ws
set acceptedKeys to {}
set rejectedKeys to {}
set has title of chart of obj to true
set chart title text of chart title of chart of obj to "Updated native revenue"
set end of acceptedKeys to "chart_title"
set chart type of chart of obj to column clustered
set end of acceptedKeys to "chart_type"
set left position of obj to 30
set end of acceptedKeys to "geometry.left"
set width of obj to 420
set end of acceptedKeys to "geometry.width"
return "{" & "\"accepted\":" & my j(acceptedKeys) & ",\"rejected\":" & my j(rejectedKeys) & "}"
 end tell
end timeout
