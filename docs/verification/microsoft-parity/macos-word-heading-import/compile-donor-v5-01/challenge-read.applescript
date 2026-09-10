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
set sourceStyle to Word style (style heading1) of boundDoc
set cloneStyle to Word style "WPSC Heading Clone" of boundDoc
set nativeRows to {{"challenge",font size of font object of sourceStyle,space before of paragraph format of sourceStyle,font size of font object of text object of paragraph 1 of boundDoc,space before of paragraph format of text object of paragraph 1 of boundDoc,font size of font object of cloneStyle,space before of paragraph format of cloneStyle,font size of font object of text object of paragraph 4 of boundDoc,space before of paragraph format of text object of paragraph 4 of boundDoc}}
end tell
return my jsonRows(nativeRows)
