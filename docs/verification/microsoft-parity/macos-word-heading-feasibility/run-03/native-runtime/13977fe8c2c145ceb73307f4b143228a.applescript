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
set boundDoc to document "document-5b965b91c2e34e0cb485a76e98371261.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-r552dodq/document-5b965b91c2e34e0cb485a76e98371261.docx" then error "WPSC_STALE_DOCUMENT"
set sentinelDoc to make new document
set content of text object of sentinelDoc to "HEADING ORDER SENTINEL e6c22adc1351472597635716bf188fc9 中文😀"
set nativeRows to {{name of sentinelDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
