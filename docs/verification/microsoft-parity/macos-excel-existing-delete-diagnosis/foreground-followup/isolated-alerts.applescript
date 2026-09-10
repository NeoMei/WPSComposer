tell application "/Applications/Microsoft Excel.app"
if name of every workbook is not {"isolated-launch.xlsx"} then error "wrong inventory"
set display alerts to false
return display alerts
end tell
