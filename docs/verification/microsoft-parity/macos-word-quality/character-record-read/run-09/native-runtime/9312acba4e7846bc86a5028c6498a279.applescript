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
set boundDoc to document "document-a222498c35e141d9a33a3f8f51902e3e.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mlaabxvv/document-a222498c35e141d9a33a3f8f51902e3e.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mlaabxvv/document-a222498c35e141d9a33a3f8f51902e3e.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mlaabxvv/document-a222498c35e141d9a33a3f8f51902e3e.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mlaabxvv/document-a222498c35e141d9a33a3f8f51902e3e.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "full-format-records-1:0: repeat with probePosition from 0 to 219"
log probePhase
repeat with probePosition from 0 to 219
set probePhase to "full-format-records-1:1: set probeRange to create range boundDoc start probePosition end (probePosition+1)"
log probePhase
set probeRange to create range boundDoc start probePosition end (probePosition+1)
set probePhase to "full-format-records-1:2: set probeFontRecord to (get properties of font object of probeRange) as record"
log probePhase
set probeFontRecord to (get properties of font object of probeRange) as record
set probePhase to "full-format-records-1:3: set probeShadingRecord to (get properties of shading of probeRange) as record"
log probePhase
set probeShadingRecord to (get properties of shading of probeRange) as record
set probePhase to "full-format-records-1:4: set end of probeRows to {probePosition as integer,name of probeFontRecord as text,font size of probeFontRecord,bold of probeFontRecord,itali"
log probePhase
set end of probeRows to {probePosition as integer,name of probeFontRecord as text,font size of probeFontRecord,bold of probeFontRecord,italic of probeFontRecord,underline of probeFontRecord as text,color of probeFontRecord,background pattern color of probeShadingRecord}
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
