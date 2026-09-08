with timeout of 30 seconds
tell application "/Applications/Microsoft Excel.app"
set b to make new workbook
set value of range "A1" of worksheet 1 of b to "unsaved-session-final-74e4"
return name of b
end tell
end timeout