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
set boundDoc to document "document-24b9eeeefdfa401d98aeaf110636b092.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-tmtm7mzf/document-24b9eeeefdfa401d98aeaf110636b092.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-tmtm7mzf/document-24b9eeeefdfa401d98aeaf110636b092.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-tmtm7mzf/document-24b9eeeefdfa401d98aeaf110636b092.docx") then error "QUALITY_WINDOW_CHANGED"
set qr to text object of selection of boundWindow
set qa to text object of selection of active window
set qb to bookmark "wpsc_document_quality_anchor" of boundDoc
set nativeRows to {{"reservation",posix full name of document of selection of boundWindow as text,name of document of active window as text,start of content of qr,end of content of qr,start of content of qa,end of content of qa,start of bookmark of qb,end of bookmark of qb,empty of qb}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
