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
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"SOURCE APPEARANCE 中文 العربية 😀" & return & "REPLACE" & return & "SUFFIX" & return & "" & return & "") as boolean) then error "WPSC_COLLAPSED_FULL_PREIMAGE"
set importRange to create range boundDoc start 32 end 32
set importStart to start of content of importRange
set importEnd to end of content of importRange
if importStart is not 32 or importEnd is not 32 then error "WPSC_COLLAPSED_RANGE"
insert file at importRange file name "/private/donor.docx" confirm conversions false link false
set nativeRows to {{"insert",importStart,importEnd,content of text object of boundDoc as text}}
end tell
return my jsonRows(nativeRows)
