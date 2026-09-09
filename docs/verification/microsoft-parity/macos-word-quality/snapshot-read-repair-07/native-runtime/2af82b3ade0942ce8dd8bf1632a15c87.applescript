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
set boundDoc to document "document-64ca299027e84fffa416ef4228f09746.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-kd7_754d/document-64ca299027e84fffa416ef4228f09746.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-kd7_754d/document-64ca299027e84fffa416ef4228f09746.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-kd7_754d/document-64ca299027e84fffa416ef4228f09746.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-kd7_754d/document-64ca299027e84fffa416ef4228f09746.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "bulk-shading:0: set probeValues to \"unresolved-sentinel\""
log probePhase
set probeValues to "unresolved-sentinel"
set probePhase to "bulk-shading:1: set probeValues to get background pattern color of shading of every character of text object of boundDoc"
log probePhase
set probeValues to get background pattern color of shading of every character of text object of boundDoc
set probePhase to "bulk-shading:2: set probeVectorClass to (class of probeValues) as text"
log probePhase
set probeVectorClass to (class of probeValues) as text
set probePhase to "bulk-shading:3: if class of probeValues is not list then error \"BULK_VALUES_NOT_LIST\""
log probePhase
if class of probeValues is not list then error "BULK_VALUES_NOT_LIST"
set probePhase to "bulk-shading:4: set probeScalarValues to {}"
log probePhase
set probeScalarValues to {}
set probePhase to "bulk-shading:5: repeat with probeValueOrdinal from 1 to count probeValues"
log probePhase
repeat with probeValueOrdinal from 1 to count probeValues
set probePhase to "bulk-shading:6: set probeValue to item probeValueOrdinal of probeValues"
log probePhase
set probeValue to item probeValueOrdinal of probeValues
set probePhase to "bulk-shading:7: if class of probeValue is not list then error \"BULK_RGB_NOT_LIST\""
log probePhase
if class of probeValue is not list then error "BULK_RGB_NOT_LIST"
set probePhase to "bulk-shading:8: set probeRgbValues to {}"
log probePhase
set probeRgbValues to {}
set probePhase to "bulk-shading:9: repeat with probeRgbOrdinal from 1 to count probeValue"
log probePhase
repeat with probeRgbOrdinal from 1 to count probeValue
set probePhase to "bulk-shading:10: set end of probeRgbValues to item probeRgbOrdinal of probeValue as integer"
log probePhase
set end of probeRgbValues to item probeRgbOrdinal of probeValue as integer
end repeat
set probePhase to "bulk-shading:12: set end of probeScalarValues to probeRgbValues"
log probePhase
set end of probeScalarValues to probeRgbValues
end repeat
set probePhase to "bulk-shading:14: set probeRows to {{\"vector\",probeVectorClass,(count probeScalarValues) as integer,probeScalarValues}}"
log probePhase
set probeRows to {{"vector",probeVectorClass,(count probeScalarValues) as integer,probeScalarValues}}
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
