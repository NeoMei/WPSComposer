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
set boundDoc to document "document-928d27e1cc144271a6a764b030886d82.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-fk90_aps/document-928d27e1cc144271a6a764b030886d82.docx" then error "WPSC_STALE_DOCUMENT"
set requestedHeading1 to Word style (style heading1) of boundDoc
set requestedHeading2 to Word style (style heading2) of boundDoc
set requestedHeading3 to Word style (style heading3) of boundDoc
set requestedHeading4 to Word style (style heading4) of boundDoc
set requestedHeading5 to Word style (style heading5) of boundDoc
set requestedHeading6 to Word style (style heading6) of boundDoc
set color of font object of requestedHeading1 to {8738, 17476, 26214}
set color of font object of requestedHeading2 to {8738, 17476, 26214}
set color of font object of requestedHeading3 to {8738, 17476, 26214}
set color of font object of requestedHeading4 to {8738, 17476, 26214}
set color of font object of requestedHeading5 to {8738, 17476, 26214}
set color of font object of requestedHeading6 to {8738, 17476, 26214}
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
