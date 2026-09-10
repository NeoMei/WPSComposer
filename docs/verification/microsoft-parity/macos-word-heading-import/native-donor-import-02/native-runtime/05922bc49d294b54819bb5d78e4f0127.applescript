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
with timeout of 60 seconds
tell application "/Applications/Microsoft Word.app"
set nativeRows to {}
set boundDoc to document "document-794a3014b43a4fa0a3ec5f7c9f84909c.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-b09vsv47/document-794a3014b43a4fa0a3ec5f7c9f84909c.docx" then error "WPSC_STALE_DOCUMENT"
set detachedStyle to Word style "WPSC Heading Clone" of boundDoc
set p to (end of content of text object of boundDoc) - 1
set freshRange to create range boundDoc start p end p
set content of freshRange to "FRESH STYLE 中文 العربية 😀" & return
set freshRange to text object of paragraph 4 of boundDoc
set style of freshRange to detachedStyle
reset font object of freshRange
reset paragraph format of freshRange
set nativeRows to {{"stage",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
