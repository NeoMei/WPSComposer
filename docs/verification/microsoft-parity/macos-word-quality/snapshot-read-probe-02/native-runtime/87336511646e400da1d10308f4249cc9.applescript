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
set boundDoc to document "document-0c387546be794533b90eb09d73ac5b59.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-oqfpzl08/document-0c387546be794533b90eb09d73ac5b59.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-oqfpzl08/document-0c387546be794533b90eb09d73ac5b59.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-oqfpzl08/document-0c387546be794533b90eb09d73ac5b59.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-oqfpzl08/document-0c387546be794533b90eb09d73ac5b59.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "table-1-ordinal-rows:0: set qt to table 1 of boundDoc"
log probePhase
set qt to table 1 of boundDoc
set probePhase to "table-1-ordinal-rows:1: repeat with probeRowOrdinal from 1 to count rows of qt"
log probePhase
repeat with probeRowOrdinal from 1 to count rows of qt
set probePhase to "table-1-ordinal-rows:2: set qrow to row probeRowOrdinal of qt"
log probePhase
set qrow to row probeRowOrdinal of qt
set probePhase to "table-1-ordinal-rows:3: set probeFlag to allow break across pages of qrow"
log probePhase
set probeFlag to allow break across pages of qrow
set probePhase to "table-1-ordinal-rows:4: set end of probeRows to {\"row\",probeRowOrdinal as integer,probeFlag}"
log probePhase
set end of probeRows to {"row",probeRowOrdinal as integer,probeFlag}
end repeat
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
