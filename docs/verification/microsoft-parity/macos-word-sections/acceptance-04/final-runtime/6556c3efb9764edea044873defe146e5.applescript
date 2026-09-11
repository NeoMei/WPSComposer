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
set boundDoc to document "document-27c4214ea733412aa6547778a3411805.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-pcfxqhp2/document-27c4214ea733412aa6547778a3411805.docx" then error "WPSC_STALE_DOCUMENT"
set sentinelDoc to document "文档126"
if (content of text object of sentinelDoc as text) is not "Section sentinel 881dc777a6714a46aa972ea0f9f7e1ab 中文😀" & return then error "SENTINEL_CHANGED"
if saved of sentinelDoc then error "SENTINEL_SAVED"
close sentinelDoc saving no
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
