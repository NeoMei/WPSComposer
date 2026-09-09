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
set boundDoc to document "document-f6a26a40df8a43cc8bce00276d1790c7.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-0wqhf_fp/document-f6a26a40df8a43cc8bce00276d1790c7.docx" then error "WPSC_STALE_DOCUMENT"
set r to create range boundDoc start 0 end 8
make new bookmark at boundDoc with properties {name:"wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa",text object:r}
set nativeRows to {{"seed-ack"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
