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
with timeout of 30 seconds
tell application "/Applications/Microsoft Word.app"
set boundDoc to document "owned.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/COMPILE_ONLY/owned.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/COMPILE_ONLY/owned.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/COMPILE_ONLY/owned.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "COMPILE_ONLY_SENTINEL"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"COMPILE_ONLY_SENTINEL") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"COMPILE_ONLY_TOKEN" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
activate object boundWindow
set selection start of selection of boundWindow to 12
set selection end of selection of boundWindow to 19
set diagActiveOwned to ((current application's NSString's stringWithString:(posix full name of document of active window as text))'s isEqualToString:"/COMPILE_ONLY/owned.docx") as boolean
if diagActiveOwned is not true then error "DIAG_ACTIVE_OWNER_CHANGED"
set diagSelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of document of diagSelection as text))'s isEqualToString:"/COMPILE_ONLY/owned.docx") then error "DIAG_SELECTION_OWNER_CHANGED"
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
set diagTable to make new table at boundDoc with properties {text object:diagRange,number of rows:4,number of columns:3}
on error diagCaughtMessage number diagCaughtCode
if diagCaughtCode is not -2710 then error diagCaughtMessage number diagCaughtCode
set diagResult to "ordinary-error"
set diagCode to diagCaughtCode
set diagMessage to diagCaughtMessage
end try
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/COMPILE_ONLY/owned.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/COMPILE_ONLY/owned.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "COMPILE_ONLY_SENTINEL"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"COMPILE_ONLY_SENTINEL") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"COMPILE_ONLY_TOKEN" & return & "") then error "QUALITY_SENTINEL_TEXT"
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
