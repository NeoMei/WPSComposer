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
set boundDoc to document "document-9e86291357fc4fb19f4d99fe733ccf38.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-d9_1w7yh/document-9e86291357fc4fb19f4d99fe733ccf38.docx" then error "WPSC_STALE_DOCUMENT"
set currentBound to (end of content of text object of boundDoc) - 1
set rollbackRange to create range boundDoc start 67 end currentBound
set content of rollbackRange to ""
set nativeRows to {{"content-clear-ack",(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
