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
set probePhase to "record-vs-legacy-20-30:0: set probeFirst to 20"
log probePhase
set probeFirst to 20
set probePhase to "record-vs-legacy-20-30:1: set probeLast to 29"
log probePhase
set probeLast to 29
set probePhase to "record-vs-legacy-20-30:2: repeat with probePosition from probeFirst to probeLast"
log probePhase
repeat with probePosition from probeFirst to probeLast
set probePhase to "record-vs-legacy-20-30:3: set probeRange to create range boundDoc start probePosition end (probePosition+1)"
log probePhase
set probeRange to create range boundDoc start probePosition end (probePosition+1)
set probePhase to "record-vs-legacy-20-30:4: set probeFontReference to font object of probeRange"
log probePhase
set probeFontReference to font object of probeRange
set probePhase to "record-vs-legacy-20-30:5: set probeLegacyRow to {name of probeFontReference as text,font size of probeFontReference,bold of probeFontReference,italic of probeFontRefe"
log probePhase
set probeLegacyRow to {name of probeFontReference as text,font size of probeFontReference,bold of probeFontReference,italic of probeFontReference,underline of probeFontReference as text,color of probeFontReference,background pattern color of shading of probeRange}
set probePhase to "record-vs-legacy-20-30:6: set probeFontRecord to (get properties of font object of probeRange) as record"
log probePhase
set probeFontRecord to (get properties of font object of probeRange) as record
set probePhase to "record-vs-legacy-20-30:7: set probeShadingRecord to (get properties of shading of probeRange) as record"
log probePhase
set probeShadingRecord to (get properties of shading of probeRange) as record
set probePhase to "record-vs-legacy-20-30:8: set probeRecordRow to {name of probeFontRecord as text,font size of probeFontRecord,bold of probeFontRecord,italic of probeFontRecord,underl"
log probePhase
set probeRecordRow to {name of probeFontRecord as text,font size of probeFontRecord,bold of probeFontRecord,italic of probeFontRecord,underline of probeFontRecord as text,color of probeFontRecord,background pattern color of probeShadingRecord}
set probePhase to "record-vs-legacy-20-30:9: set probeLegacyJson to my jsonRows({probeLegacyRow})"
log probePhase
set probeLegacyJson to my jsonRows({probeLegacyRow})
set probePhase to "record-vs-legacy-20-30:10: set probeRecordJson to my jsonRows({probeRecordRow})"
log probePhase
set probeRecordJson to my jsonRows({probeRecordRow})
set probePhase to "record-vs-legacy-20-30:11: set probeEqual to (current application's NSString's stringWithString:probeLegacyJson)'s isEqualToString:probeRecordJson"
log probePhase
set probeEqual to (current application's NSString's stringWithString:probeLegacyJson)'s isEqualToString:probeRecordJson
set probePhase to "record-vs-legacy-20-30:12: if not (probeEqual as boolean) then error \"RECORD_VALUE_MISMATCH\""
log probePhase
if not (probeEqual as boolean) then error "RECORD_VALUE_MISMATCH"
set probePhase to "record-vs-legacy-20-30:13: set end of probeRows to {\"same-coordinate\",probePosition as integer,probeEqual as boolean}"
log probePhase
set end of probeRows to {"same-coordinate",probePosition as integer,probeEqual as boolean}
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
