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
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"" & return & "REPLACE" & return & "SUFFIX" & return & "" & return & "") as boolean) then error "WPSC_COLLAPSED_FULL_PREIMAGE"
set carrierRange to text object of paragraph 1 of boundDoc
if start of content of carrierRange is not 0 or end of content of carrierRange is not 1 or content of carrierRange is not return then error "WPSC_V3_EMPTY_CARRIER"
set controlStyle to Word style (style heading1) of boundDoc
set style of carrierRange to controlStyle
set freshInsert to create range boundDoc start 0 end 0
set content of freshInsert to "FRESH STYLE 中文 العربية 😀"
set nativeRows to {{"stage",true}}
end tell
return my jsonRows(nativeRows)
