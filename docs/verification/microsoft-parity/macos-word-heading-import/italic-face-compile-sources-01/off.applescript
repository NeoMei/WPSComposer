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
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_ITALIC_BODY"
set sourceStyle to Word style (style heading1) of boundDoc
set linkedStyle to Word style "标题 1 字符" of boundDoc
if italic of font object of sourceStyle is not true then error "WPSC_ITALIC_PREIMAGE"
if italic of font object of linkedStyle is not true then error "WPSC_ITALIC_PAIR_PREIMAGE"
set italic of font object of sourceStyle to false
set nativeRows to {{"italic",italic of font object of sourceStyle,italic of font object of linkedStyle}}
end tell
return my jsonRows(nativeRows)
