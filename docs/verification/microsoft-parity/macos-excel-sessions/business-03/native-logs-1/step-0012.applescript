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
set ownedBook to workbook "owned-46d20261a45b411ab4486cd7f4e24191.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-uk1hn99_/owned-46d20261a45b411ab4486cd7f4e24191.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 1 of ownedBook
activate object ws
set obj to range "B2" of ws
set ownedWindow to window 1 of ownedBook
set freeze panes of ownedWindow to false
set split column of ownedWindow to (first column index of obj) - 1
set split row of ownedWindow to (first row index of obj) - 1
select obj
set freeze panes of ownedWindow to true
return "{" & "\"freeze_panes\":" & my j(freeze panes of ownedWindow) & ",\"split_column\":" & my j(split column of ownedWindow) & ",\"split_row\":" & my j(split row of ownedWindow) & "}"
 end tell
end timeout
