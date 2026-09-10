tell application "/Applications/Microsoft Excel.app"
set out to {}
repeat with b in workbooks
set end of out to {name of b, full name of b, saved of b, value of range "A1" of worksheet 1 of b}
end repeat
return {display alerts, out}
end tell
