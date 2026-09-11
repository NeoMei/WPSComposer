tell application "/Applications/Microsoft Excel.app"
if name of every workbook is not {"owned-ec3f072bb4ed4252924818c9e68dfbf4.xlsx"} then error "wrong inventory"
set wb to workbook 1
if full name of wb is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-cvcue6ha/owned-ec3f072bb4ed4252924818c9e68dfbf4.xlsx" then error "wrong path"
if name of every worksheet of wb is not {"Data","Results"} then error "wrong sheets"
if value of range "Z40" of worksheet 1 of wb is not "PRIVATE-JSON-ROUNDTRIP" then error "wrong marker"
if saved of wb is not true then error "unsaved"
if display alerts is not true then error "alerts changed"
close wb saving no
if (count of workbooks) is not 0 then error "not closed"
quit saving no
return "CLEANUP_PASS"
end tell
