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
set boundDoc to document "document-7d9a01c7df424a45beddaf520fc4b39c.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-kyut87st/document-7d9a01c7df424a45beddaf520fc4b39c.docx" then error "WPSC_STALE_DOCUMENT"
set probeRange to create range boundDoc start 11 end 41
set nativeRows to {{start of content of probeRange,end of content of probeRange,content of probeRange as text,italic of font object of probeRange,(color of font object of probeRange is {40092, 0, 1542}),(background pattern color of shading of probeRange is {64764, 59624, 59110})}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
