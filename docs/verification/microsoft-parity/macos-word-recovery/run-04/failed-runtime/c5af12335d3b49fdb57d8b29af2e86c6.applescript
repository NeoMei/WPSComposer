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
set boundDoc to document "document-210ca8985c0942949d66360c6de56031.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-zqca0bbu/document-210ca8985c0942949d66360c6de56031.docx" then error "WPSC_STALE_DOCUMENT"
set d to document "文档2"
if (content of text object of d as text) is not "Recovery sentinel 中文😀 a1678b24c53d49629499aa515a8087bf" & return then error "SENTINEL_CHANGED"
if saved of d then error "SENTINEL_SAVED"
close d saving no
set nativeRows to {{"closed"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
