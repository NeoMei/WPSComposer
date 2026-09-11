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
set boundDoc to document "document-be6dfc66b0844d34a50444209ac0b5c5.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-7tfd40uw/document-be6dfc66b0844d34a50444209ac0b5c5.docx" then error "WPSC_STALE_DOCUMENT"
set figureSentinel to make new document
set content of text object of figureSentinel to "FIGURE ROLLBACK SENTINEL 中文😀 987e7474828e457e9f05e4b2fd4b3c9e"
set figureSentinelWindow to active window of figureSentinel
activate object figureSentinelWindow
set nativeRows to {{name of figureSentinel as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
