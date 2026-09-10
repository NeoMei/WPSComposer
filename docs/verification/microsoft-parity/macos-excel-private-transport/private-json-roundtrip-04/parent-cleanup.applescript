tell application "/Applications/Microsoft Excel.app"
if (count of workbooks) is not 1 then error "foreign inventory"
set ownedBook to workbook 1
if full name of ownedBook is not "/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/session-86z9pg58/owned-80bd96fb948f4cf3a16961d791a300ff.xlsx" then error "path mismatch"
if name of every worksheet of ownedBook is not {"Data","Results"} then error "sheet mismatch"
if value of range "Z40" of worksheet 1 of ownedBook is not "PRIVATE-JSON-ROUNDTRIP" then error "marker mismatch"
if saved of ownedBook is not true then error "unexpected dirty state"
if display alerts is not true then error "alerts changed"
close ownedBook saving no
if (count of workbooks) is not 0 then error "close not acknowledged"
quit saving no
return "OWNED_CLEANUP_PASS"
end tell
