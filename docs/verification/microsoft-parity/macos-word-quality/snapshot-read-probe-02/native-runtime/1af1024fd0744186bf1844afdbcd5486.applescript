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
set probePhase to "style-named-full-properties:0: set qualityStyleNames to get name local of every Word style of boundDoc"
log probePhase
set qualityStyleNames to get name local of every Word style of boundDoc
set probePhase to "style-named-full-properties:1: repeat with qualityStyleOrdinal from 1 to count qualityStyleNames"
log probePhase
repeat with qualityStyleOrdinal from 1 to count qualityStyleNames
set probePhase to "style-named-full-properties:2: set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text"
log probePhase
set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text
set probePhase to "style-named-full-properties:3: set qstyle to Word style qualityStyleName of boundDoc"
log probePhase
set qstyle to Word style qualityStyleName of boundDoc
set probePhase to "style-named-full-properties:4: set end of probeRows to {\"style\",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}"
log probePhase
set end of probeRows to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}
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
