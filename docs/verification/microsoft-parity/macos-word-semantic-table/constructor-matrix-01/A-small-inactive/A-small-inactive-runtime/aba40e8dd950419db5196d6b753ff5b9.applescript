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
set boundDoc to document "document-e2218bec0cec4030885c48597afb9b47.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx") then error "QUALITY_WINDOW_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "文档92"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档92") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"SEMANTIC MATRIX SENTINEL 中文😀 da007ee4a6e74be6b7a5e39b76e7c136" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
set selection start of selection of boundWindow to 12
set selection end of selection of boundWindow to 19
activate object (active window of qualitySentinel)
set diagActiveOwned to ((current application's NSString's stringWithString:(posix full name of document of active window as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx") as boolean
if diagActiveOwned is not false then error "DIAG_ACTIVE_OWNER_CHANGED"
set diagSelection to selection of boundWindow
if story type of diagSelection is not main text story then error "DIAG_STORY_CHANGED"
set diagRange to text object of diagSelection
if start of content of diagRange is not 12 or end of content of diagRange is not 19 then error "DIAG_RANGE_CHANGED"
if (count tables of boundDoc) is not 1 then error "DIAG_SEED_CHANGED"
set diagPrefix to create range boundDoc start 0 end 12
set diagFullMarker to create range boundDoc start 12 end 19
if (content of diagFullMarker as text) is not "REPLACE" then error "DIAG_MARKER_CHANGED"
set nativeRows to {{"before",posix full name of boundDoc as text,12,end of content of diagRange,count tables of boundDoc,content of text object of boundDoc as text,content of text object of bookmark "semantic_old_table" of boundDoc as text,content of diagPrefix as text,content of text object of bookmark "semantic_suffix" of boundDoc as text,saved of boundDoc,name of document of active window as text,diagActiveOwned}}
set diagResult to "created"
set diagCode to 0
set diagMessage to ""
try
set diagTable to make new table at boundDoc with properties {text object:diagRange,number of rows:1,number of columns:1}
on error diagMessage number diagCode
if diagCode is not -2710 then error diagMessage number diagCode
set diagResult to "ordinary-error"
end try
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-9fkkpaxe/document-e2218bec0cec4030885c48597afb9b47.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "文档92"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档92") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"SEMANTIC MATRIX SENTINEL 中文😀 da007ee4a6e74be6b7a5e39b76e7c136" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
set end of nativeRows to {"attempt",diagResult,diagCode,diagMessage}
set end of nativeRows to {"after",count tables of boundDoc,content of text object of boundDoc as text,saved of boundDoc}
repeat with diagIndex from 1 to (count tables of boundDoc)
set diagObserved to table diagIndex of boundDoc
set end of nativeRows to {"table",diagIndex,start of content of text object of diagObserved,end of content of text object of diagObserved,number of rows of diagObserved,number of columns of diagObserved,content of text object of diagObserved as text}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
