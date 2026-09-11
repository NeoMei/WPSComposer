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
set boundDoc to document "document.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mex9ycx0/document.docx" then error "WPSC_STALE_DOCUMENT"
set targetRange to text object of paragraph 1 of boundDoc
set sourceRange to targetRange
set sourceStart to start of content of sourceRange
set sourceEnd to end of content of sourceRange
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint >= sourceStart and insertionPoint <= sourceEnd then error "WPSC_OVERLAPPING_RANGE"
set insertionRange to create range boundDoc start insertionPoint end insertionPoint
set formatted text of insertionRange to formatted text of sourceRange
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows(nativeRows)
