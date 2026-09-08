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
set boundDoc to document "document-e00a0b258a2c423eb0bdea49174c7cb9.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-sw_iddow/document-e00a0b258a2c423eb0bdea49174c7cb9.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {}
repeat with documentIndex from 1 to (count documents)
set inventoryDoc to document documentIndex
if (posix full name of inventoryDoc as text) is not (posix full name of boundDoc as text) then
set end of nativeRows to {name of inventoryDoc as text, posix full name of inventoryDoc as text, saved of inventoryDoc, content of text object of inventoryDoc as text}
end if
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
