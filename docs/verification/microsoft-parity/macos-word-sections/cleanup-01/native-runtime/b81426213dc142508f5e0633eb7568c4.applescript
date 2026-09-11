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
set boundDoc to document "document-bf1a9ebaee0f4ff893c2082fccd2e309.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-3yo7qjie/document-bf1a9ebaee0f4ff893c2082fccd2e309.docx" then error "WPSC_STALE_DOCUMENT"
set sentinelDoc to document "文档114"
if (posix full name of sentinelDoc as text) is not "文档114" then error "SENTINEL_PATH_CHANGED"
if saved of sentinelDoc then error "SENTINEL_SAVED"
if (content of text object of sentinelDoc as text) is not "Section sentinel 4d43efabfbcb410b8528aa011edc0fd5 中文😀" & return & "" then error "SENTINEL_TEXT_CHANGED"
close sentinelDoc saving no
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
