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
set sourceStyle to Word style (style heading1) of boundDoc
set linkedStyle to Word style "标题 1 字符" of boundDoc
set cloneStyle to Word style "WPSC Heading Clone" of boundDoc
set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc,italic of font object of sourceStyle,italic of font object of linkedStyle,italic of font object of cloneStyle}}
end tell
return my jsonRows(nativeRows)
