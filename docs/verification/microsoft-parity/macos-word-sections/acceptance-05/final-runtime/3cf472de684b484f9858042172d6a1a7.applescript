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
set boundDoc to document "document-630ad4e46b324d08869b4fab94ea7331.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-bu3wg936/document-630ad4e46b324d08869b4fab94ea7331.docx" then error "WPSC_STALE_DOCUMENT"
set sentinelDoc to document "文档128"
if (content of text object of sentinelDoc as text) is not "Section sentinel 76acdbf75e2e4f50b5fc29118e16d4ab 中文😀" & return then error "SENTINEL_CHANGED"
if saved of sentinelDoc then error "SENTINEL_SAVED"
close sentinelDoc saving no
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
