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
-- WPSC_BIND_OPEN
open file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-fm4cuket/document-44f9b7c7bb234fd1be8a71dcfe1f95a6.docx" read only false add to recent files false
set matches to {}
repeat with di from 1 to (count of documents)
set d to document di
if (posix full name of d as text) is "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-fm4cuket/document-44f9b7c7bb234fd1be8a71dcfe1f95a6.docx" then set end of matches to d
end repeat
if (count matches) is not 1 then error "WPSC_BINDING_FAILED"
set boundDoc to item 1 of matches
set boundWindow to active window of boundDoc
set wid to id of boundWindow
set nativeRows to {{"binding", wid, posix full name of boundDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
