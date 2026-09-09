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
set boundDoc to document "document-a22d1c02449b434b9d467f2e8808f9be.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-bdwhv31_/document-a22d1c02449b434b9d467f2e8808f9be.docx" then error "WPSC_STALE_DOCUMENT"
set currentBound to (end of content of text object of boundDoc) - 1
set targetBound to 0
if targetBound < 0 or targetBound > currentBound then error "WPSC_INVALID_ROLLBACK_BOUND"
if targetBound < currentBound then
set doomedRange to create range boundDoc start targetBound end currentBound
delete doomedRange
end if
set nativeRows to {{"rollback",targetBound,(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
