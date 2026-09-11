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
set boundDoc to document "document-1ae4884692704bb39a15dc26a88979a4.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "primary-header-evaluated-content:0: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "primary-header-evaluated-content:1: set probeStoryRange to get story range boundDoc story type primary header story"
log probePhase
set probeStoryRange to get story range boundDoc story type primary header story
set probePhase to "primary-header-evaluated-content:2: set probeStoryRange to get probeStoryRange"
log probePhase
set probeStoryRange to get probeStoryRange
set probePhase to "primary-header-evaluated-content:3: set probeValue to \"unresolved-sentinel\""
log probePhase
set probeValue to "unresolved-sentinel"
set probePhase to "primary-header-evaluated-content:4: set probeValue to get content of probeStoryRange"
log probePhase
set probeValue to get content of probeStoryRange
set probePhase to "primary-header-evaluated-content:5: set probeValueClass to (class of probeValue) as text"
log probePhase
set probeValueClass to (class of probeValue) as text
set probePhase to "primary-header-evaluated-content:6: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "primary-header-evaluated-content:7: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "primary-header-evaluated-content:8: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "primary-header-evaluated-content:9: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "primary-header-evaluated-content:10: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "primary-header-evaluated-content:11: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "primary-header-evaluated-content:12: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "primary-header-evaluated-content:13: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "primary-header-evaluated-content:14: set recoveryData to (current application's NSString's stringWithString:(probeValue as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(probeValue as text))'s dataUsingEncoding:4
set probePhase to "primary-header-evaluated-content:15: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "primary-header-evaluated-content:16: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "primary-header-evaluated-content:17: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "primary-header-evaluated-content:18: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "primary-header-evaluated-content:19: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "primary-header-evaluated-content:20: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "primary-header-evaluated-content:21: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "primary-header-evaluated-content:22: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "primary-header-evaluated-content:23: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "primary-header-evaluated-content:24: set probeValueHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set probeValueHash to text 1 thru 64 of recoveryDigestText
set probePhase to "primary-header-evaluated-content:25: set probeRows to {{\"value-hash\",probeValueClass,probeValueHash}}"
log probePhase
set probeRows to {{"value-hash",probeValueClass,probeValueHash}}
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
