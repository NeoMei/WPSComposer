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
set boundDoc to document "document-77a00d3c690d42fcb60cd1733d903bea.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-e2k8a_fg/document-77a00d3c690d42fcb60cd1733d903bea.docx" then error "WPSC_STALE_DOCUMENT"
set sentinelDoc to make new document
set content of text object of sentinelDoc to "WPSC-OBJECT-SENTINEL-9f50cd29489a41aab9399e298683d14f"
set nativeRows to {{"sentinel", name of sentinelDoc as text, posix full name of sentinelDoc as text, saved of sentinelDoc, content of text object of sentinelDoc as text, id of active window of sentinelDoc}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
