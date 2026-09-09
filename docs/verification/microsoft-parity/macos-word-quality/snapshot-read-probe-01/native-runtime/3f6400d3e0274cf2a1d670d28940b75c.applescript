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
set boundDoc to document "document-6324ad13159642c1ab0ce315541c35b0.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-crg8k8ws/document-6324ad13159642c1ab0ce315541c35b0.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-crg8k8ws/document-6324ad13159642c1ab0ce315541c35b0.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-crg8k8ws/document-6324ad13159642c1ab0ce315541c35b0.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-crg8k8ws/document-6324ad13159642c1ab0ce315541c35b0.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set qualityPrefixHash to ""
set qualitySuffixHash to ""
set probePhase to "start"
try
repeat with qstyle in Word styles of boundDoc
set probePhase to "1: set end of qualityLayout to {\"style\",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}"
set end of qualityLayout to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}
end repeat
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probeHash to text 1 thru 64 of recoveryDigestText
set nativeRows to {{"read-ok",count qualityLayout,probeHash}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
