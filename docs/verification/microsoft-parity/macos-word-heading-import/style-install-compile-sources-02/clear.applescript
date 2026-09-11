use framework "Foundation"
use scripting additions
on jsonRows(rows)
  set dataValue to current application's NSJSONSerialization's dataWithJSONObject:rows options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:dataValue encoding:4) as text
end jsonRows
on enumIndex(v, choices)
  repeat with i from 1 to count choices
    if v is item i of choices then return i - 1
  end repeat
  return -1
end enumIndex

tell application "Microsoft Word"
set boundDoc to document "COMPILE_ONLY.docx"
set boundWindow to active window of boundDoc
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "WPSC INSTALL CARRIER 中文😀" & return & "SUFFIX" & return & "FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_COLLAPSED_FULL_PREIMAGE"
set carrierRange to create range boundDoc start 60 end 86
if content of carrierRange is not "WPSC INSTALL CARRIER 中文😀" & return then error "WPSC_CARRIER_TEXT"
if start of content of carrierRange is not 60 or end of content of carrierRange is not 86 then error "WPSC_CARRIER_BOUNDS"
set content of carrierRange to ""
set ownSelection to selection of boundWindow
set selection start of ownSelection to 0
set selection end of ownSelection to 5
set nativeRows to {{"clear",content of text object of boundDoc as text}}
end tell
return my jsonRows(nativeRows)
