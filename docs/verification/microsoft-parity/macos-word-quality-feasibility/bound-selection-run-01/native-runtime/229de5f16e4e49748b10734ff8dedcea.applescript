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
set boundDoc to document "document-feb1b048cd484362b253fbbe673f5569.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx") then error "QUALITY_WINDOW_CHANGED"
activate object boundWindow
set content of text object of boundDoc to "PREFIX 中文😀" & return & "MIDDLE QUALITY ANCHOR" & return & "SUFFIX remains" & return
set selection start of selection of boundWindow to 9
set selection end of selection of boundWindow to 9
set nativeRows to {{"seed",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
