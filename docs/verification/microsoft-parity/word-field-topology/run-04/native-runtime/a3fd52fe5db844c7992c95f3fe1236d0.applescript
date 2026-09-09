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
set boundDoc to document "document-6dac3905461e4c819b4ef961232f41d1.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9cyrfp6c/document-6dac3905461e4c819b4ef961232f41d1.docx" then error "WPSC_STALE_DOCUMENT"
set r to create range boundDoc start 0 end 8
make new bookmark at boundDoc with properties {name:"wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa",text object:r}
set insertionPoint to (end of content of text object of boundDoc) - 1
set f to create range boundDoc start insertionPoint end insertionPoint
create new field text range f field type field ref field text "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa" preserve formatting true
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
