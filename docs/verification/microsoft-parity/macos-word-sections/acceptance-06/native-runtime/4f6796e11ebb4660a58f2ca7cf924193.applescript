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
set boundDoc to document "document-d0b9b67b4e86457baa963453813afbdc.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-8pn_ch2c/document-d0b9b67b4e86457baa963453813afbdc.docx" then error "WPSC_STALE_DOCUMENT"
set titleProperty to document property "Title" of boundDoc
set authorProperty to document property "Author" of boundDoc
get value of titleProperty
get value of authorProperty
set value of titleProperty to ""
if (value of titleProperty as text) is not "" then error "WPSC_METADATA_READBACK_FAILED"
set value of authorProperty to ""
if (value of authorProperty as text) is not "" then error "WPSC_METADATA_READBACK_FAILED"
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
