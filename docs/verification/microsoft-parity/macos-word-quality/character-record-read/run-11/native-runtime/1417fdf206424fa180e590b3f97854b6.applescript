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
set boundDoc to document "document-02ef2a432f834f0eb8e896153d44781a.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-nfyv8899/document-02ef2a432f834f0eb8e896153d44781a.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-nfyv8899/document-02ef2a432f834f0eb8e896153d44781a.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-nfyv8899/document-02ef2a432f834f0eb8e896153d44781a.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-nfyv8899/document-02ef2a432f834f0eb8e896153d44781a.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "production-first-format-1:0: set qa to 0"
log probePhase
set qa to 0
set probePhase to "production-first-format-1:1: set qz to 220"
log probePhase
set qz to 220
set probePhase to "production-first-format-1:2: set qualityFormats to {}"
log probePhase
set qualityFormats to {}
set probePhase to "production-first-format-1:3: repeat with qc from qa to qz - 1"
log probePhase
repeat with qc from qa to qz - 1
set probePhase to "production-first-format-1:4: set qcr to create range boundDoc start qc end (qc + 1)"
log probePhase
set qcr to create range boundDoc start qc end (qc + 1)
set probePhase to "production-first-format-1:5: set qcf to (get properties of font object of qcr) as record"
log probePhase
set qcf to (get properties of font object of qcr) as record
set probePhase to "production-first-format-1:6: set qcs to (get properties of shading of qcr) as record"
log probePhase
set qcs to (get properties of shading of qcr) as record
set probePhase to "production-first-format-1:7: set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,backgroun"
log probePhase
set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of qcs}
end repeat
set probePhase to "production-first-format-1:9: set probeRows to qualityFormats"
log probePhase
set probeRows to qualityFormats
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
