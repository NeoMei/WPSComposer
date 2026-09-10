tell application "/Applications/Microsoft Excel.app"
if (count of workbooks) is not 0 then error "not empty"
if display alerts is not true then error "alerts changed"
quit saving no
return "CLEANUP_PASS"
end tell
