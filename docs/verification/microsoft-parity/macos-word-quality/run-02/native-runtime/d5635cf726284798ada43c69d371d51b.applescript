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
set boundDoc to document "document-db70cbc45b1443f09834b808a00b9462.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-l0q6357v/document-db70cbc45b1443f09834b808a00b9462.docx") then error "QUALITY_WINDOW_CHANGED"
activate object boundWindow
set selection start of selection of boundWindow to 108
set selection end of selection of boundWindow to 111
set qualitySentinel to document "文档89"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档89") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"QUALITY STAGE ONE SENTINEL 中文😀 24f4da7430c941cf9b8cd9bcca23d081" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
activate object (active window of qualitySentinel)
set selection start of selection of active window to 2
set selection end of selection of active window to 2
set nativeRows to {{"selected",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
