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
set probePhase to "bookmarks:0: repeat with qi from 1 to count bookmarks of boundDoc"
log probePhase
repeat with qi from 1 to count bookmarks of boundDoc
set probePhase to "bookmarks:1: set qb to bookmark qi of boundDoc"
log probePhase
set qb to bookmark qi of boundDoc
set probePhase to "bookmarks:2: set qa to start of bookmark of qb"
log probePhase
set qa to start of bookmark of qb
set probePhase to "bookmarks:3: set qz to end of bookmark of qb"
log probePhase
set qz to end of bookmark of qb
set probePhase to "bookmarks:4: set qhashStart to qa"
log probePhase
set qhashStart to qa
set probePhase to "bookmarks:5: if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta"
log probePhase
if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta
set probePhase to "bookmarks:6: set qr to create range boundDoc start qhashStart end qz"
log probePhase
set qr to create range boundDoc start qhashStart end qz
set probePhase to "bookmarks:7: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "bookmarks:8: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "bookmarks:9: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "bookmarks:10: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "bookmarks:11: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "bookmarks:12: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "bookmarks:13: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "bookmarks:14: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "bookmarks:15: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "bookmarks:16: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "bookmarks:17: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "bookmarks:18: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "bookmarks:19: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "bookmarks:20: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "bookmarks:21: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "bookmarks:22: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "bookmarks:23: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "bookmarks:24: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "bookmarks:25: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "bookmarks:26: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
log probePhase
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "bookmarks:27: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
log probePhase
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "bookmarks:28: set end of qualityState to {\"bookmark\",name of qb as text,qa,qz,qualityTextHash}"
log probePhase
set end of qualityState to {"bookmark",name of qb as text,qa,qz,qualityTextHash}
end repeat
set probePhase to "bookmarks:30: set end of qualityState to {\"quality-state-end\"}"
log probePhase
set end of qualityState to {"quality-state-end"}
set nativeRows to {{"read-ok",qualityState}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
