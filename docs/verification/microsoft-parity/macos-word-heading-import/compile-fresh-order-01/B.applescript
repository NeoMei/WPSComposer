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
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"SOURCE APPEARANCE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "" & return & "") as boolean) then error "WPSC_FRESH_PREIMAGE"
set terminalRange to text object of paragraph 4 of boundDoc
if start of content of terminalRange is not 73 or content of terminalRange is not return then error "WPSC_FRESH_TERMINAL"
set detachedStyle to Word style "WPSC Heading Clone" of boundDoc
set freshRange to create range boundDoc start 73 end 73
set content of freshRange to return
set freshRange to text object of paragraph 4 of boundDoc
set style of freshRange to detachedStyle
reset font object of freshRange
reset paragraph format of freshRange
set freshStart to start of content of freshRange
set freshInsert to create range boundDoc start freshStart end freshStart
set content of freshInsert to "FRESH STYLE 中文 العربية 😀"
set nativeRows to {{"stage",true}}
end tell
return my jsonRows(nativeRows)
