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
with timeout of 61 seconds
 tell application "/Applications/Microsoft Excel.app"
set ownedBook to workbook "owned-ca0eb89ecc5245b191d1f049383475ef.xlsx"
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-bhvepely/owned-ca0eb89ecc5245b191d1f049383475ef.xlsx" then error "Owned workbook identity mismatch"
set ws to worksheet 3 of ownedBook
activate object ws
set obj to ws
if (count of workbooks) is not 1 then error "Foreign workbook in private process"
if (count of worksheets of ownedBook) <= 1 then error "Cannot remove last worksheet"
set removedSheetName to name of obj
set beforeSheetCount to count worksheets of ownedBook
set previousAlerts to display alerts
try
 set display alerts to false
 if display alerts is not false then error "Alerts not suppressed"
 delete obj
 set display alerts to previousAlerts
on error deletionMessage number deletionNumber
 set display alerts to previousAlerts
 error deletionMessage number deletionNumber
end try
if display alerts is not previousAlerts then error "Alerts not restored"
if (count worksheets of ownedBook) is not (beforeSheetCount - 1) then error "Deletion count mismatch"
if removedSheetName is in (name of every worksheet of ownedBook) then error "Deleted worksheet still present"
return "{" & "\"removed\":" & my j("sheet:3") & ",\"deleted\":" & my j(true) & "}"
 end tell
end timeout
