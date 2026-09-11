with timeout of 30 seconds
tell application "/Applications/Microsoft Excel.app"
set b to workbook "工作簿11"
if value of range "A1" of worksheet 1 of b is not "unsaved-session-final-74e4" then error "Owned marker mismatch"
if saved of b then error "Unexpected unsaved fixture state"
close b saving no
return "owned-only-close"
end tell
end timeout