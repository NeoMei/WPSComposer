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
set boundDoc to document "document.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-v_aq6plt/document.docx" then error "WPSC_STALE_DOCUMENT"
set targetRange to text object of paragraph 2 of boundDoc
set name of font object of targetRange to "Arial"
set font size of font object of targetRange to 15
set bold of font object of targetRange to true
set color of font object of targetRange to {4626, 13364, 22102}
set first line indent of paragraph format of targetRange to 24
set space after of paragraph format of targetRange to 8
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows(nativeRows)
