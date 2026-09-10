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
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_COLLAPSED_FULL_PREIMAGE"
set importRange to create range boundDoc start 60 end 60
set importStart to start of content of importRange
set importEnd to end of content of importRange
if importStart is not 60 or importEnd is not 60 then error "WPSC_COLLAPSED_RANGE"
insert file at importRange file name "/tmp/COMPILE_ONLY.docx" confirm conversions false link false
set nativeRows to {{"insert",importStart,importEnd,content of text object of boundDoc as text}}
end tell
return my jsonRows(nativeRows)
