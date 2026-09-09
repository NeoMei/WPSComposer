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
set probePhase to "font-record-59:0: set probeRange to create range boundDoc start 59 end 60"
log probePhase
set probeRange to create range boundDoc start 59 end 60
set probePhase to "font-record-59:1: set probeFontRecord to get properties of font object of probeRange"
log probePhase
set probeFontRecord to get properties of font object of probeRange
set probePhase to "font-record-59:2: set probeRecordClass to class of probeFontRecord as text"
log probePhase
set probeRecordClass to class of probeFontRecord as text
set probePhase to "font-record-59:3: set probeRows to {{\"font-record\",probeRecordClass,name of probeFontRecord as text,font size of probeFontRecord as real,bold of probeFontReco"
log probePhase
set probeRows to {{"font-record",probeRecordClass,name of probeFontRecord as text,font size of probeFontRecord as real,bold of probeFontRecord as boolean,italic of probeFontRecord as boolean,underline of probeFontRecord as text}}
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
