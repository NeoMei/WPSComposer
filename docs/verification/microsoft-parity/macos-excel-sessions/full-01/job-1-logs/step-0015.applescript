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
if full name of active workbook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-ua3u3wfv/owned-87fa307d633142aabcbf84bf86b6b695.xlsx" then error "Active selection belongs to another workbook"
set rng to selection
set rp to properties of rng
set fp to properties of font object of rng
set rangeJSON to "{" & "\"id\":" & my j("selection") & ",\"address\":" & my j(get address rng) & ",\"value\":" & my j(value of rp) & ",\"formula\":" & my j(formula of rp) & ",\"display_text\":" & my j(string value of rp) & ",\"number_format\":" & my j(number format of rp) & ",\"horizontal_alignment\":" & my j(horizontal alignment of rp) & ",\"vertical_alignment\":" & my j(vertical alignment of rp) & ",\"wrap_text\":" & my j(wrap text of rp) & ",\"indent\":" & my j(indent level of rp) & ",\"row_height\":" & my j(row height of rp) & ",\"column_width\":" & my j(column width of rp) & ",\"merged\":" & my j(merge cells of rp) & "}"
set fontJSON to "{" & "\"name\":" & my j(name of fp) & ",\"size\":" & my j(font size of fp) & ",\"bold\":" & my j(bold of fp) & ",\"italic\":" & my j(italic of fp) & ",\"underline\":" & my j(underline of fp) & ",\"strikethrough\":" & my j(strikethrough of fp) & ",\"color\":" & my j(color of fp) & "}"
set fillJSON to "{" & "\"color\":" & my j(color of interior object of rng) & "}"
set mergeJSON to "null"
if merge cells of rp is true then set mergeJSON to my j(get address merge area of rng)
set rangeJSON to (text 1 thru -2 of rangeJSON) & ",\"font\":" & fontJSON & ",\"fill\":" & fillJSON & ",\"merge_area\":" & mergeJSON & "}"

return rangeJSON
 end tell
end timeout
