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
set boundDoc to document "document-81ce5fa3221f4b7aba765048252f5f1b.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-0tgbyg94/document-81ce5fa3221f4b7aba765048252f5f1b.docx" then error "WPSC_STALE_DOCUMENT"
set sentinelDoc to document "文档215"
if (content of text object of sentinelDoc as text) is not "RULES SENTINEL 中文😀 b67e15b8d92448d895d356b593fdb324" & return then error "WPSC_SENTINEL_CHANGED"
if saved of sentinelDoc then error "WPSC_SENTINEL_SAVED"
close sentinelDoc saving no
set nativeRows to {{"sentinel",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
