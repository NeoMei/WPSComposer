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
set ownedBook to workbook "owned-ff9bede9b5544bf6bf4a9aff2e48251b.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-j1ni1yqq/owned-ff9bede9b5544bf6bf4a9aff2e48251b.xlsx" then error "Owned workbook identity mismatch"
if (count of workbooks) is not 1 then error "foreign workbook in private instance"
    if display alerts is not true then error "unexpected alerts"
    set display alerts to false
    if display alerts is not false then error "alerts not suppressed"
    try
     delete worksheet "Business report" of ownedBook
     set display alerts to true
    on error msg number n
     set display alerts to true
     error msg number n
    end try
    if name of every worksheet of ownedBook is not {"Data","Results"} then error "delete not acknowledged"
    if display alerts is not true then error "alerts not restored"
    return "{" & "\"deleted\":" & my j(true) & "}"
 end tell
end timeout
