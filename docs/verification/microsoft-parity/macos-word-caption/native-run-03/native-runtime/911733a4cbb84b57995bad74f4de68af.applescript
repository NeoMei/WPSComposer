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
set boundDoc to document "document-e58ac26ad84345a3b49ac358b9301526.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s2ci3q2b/document-e58ac26ad84345a3b49ac358b9301526.docx" then error "WPSC_STALE_DOCUMENT"
set captionSentinel to make new document
set content of text object of captionSentinel to "CAPTION SENTINEL 中文😀 bf8532b064cb41a88616cfffa47f1e9f"
set captionSentinelWindow to active window of captionSentinel
activate object captionSentinelWindow
set nativeRows to {{name of captionSentinel as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
