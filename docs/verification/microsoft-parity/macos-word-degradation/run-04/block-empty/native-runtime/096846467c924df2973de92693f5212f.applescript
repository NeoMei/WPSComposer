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
set boundDoc to document "document-cfe87eac613b48ff80872d2edf06012d.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mfq7f5dg/document-cfe87eac613b48ff80872d2edf06012d.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {}
repeat with pi from 1 to count paragraphs of boundDoc
set p to text object of paragraph pi of boundDoc
if (content of p as text) is "FOLLOWING-NORMAL-block-empty" & return then set end of nativeRows to {start of content of p,end of content of p}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
