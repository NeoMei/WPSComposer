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
set boundDoc to document "document-5f08672f6402415280dfeb7d004228fc.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-j6m3t7mi/document-5f08672f6402415280dfeb7d004228fc.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-j6m3t7mi/document-5f08672f6402415280dfeb7d004228fc.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-j6m3t7mi/document-5f08672f6402415280dfeb7d004228fc.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-j6m3t7mi/document-5f08672f6402415280dfeb7d004228fc.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "prefix-suffix-hashes:0: set qe to end of content of text object of boundDoc"
log probePhase
set qe to end of content of text object of boundDoc
set probePhase to "prefix-suffix-hashes:1: set qp to create range boundDoc start 0 end qualityPoint"
log probePhase
set qp to create range boundDoc start 0 end qualityPoint
set probePhase to "prefix-suffix-hashes:2: set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe"
log probePhase
set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe
set probePhase to "prefix-suffix-hashes:3: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "prefix-suffix-hashes:4: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "prefix-suffix-hashes:5: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "prefix-suffix-hashes:6: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "prefix-suffix-hashes:7: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "prefix-suffix-hashes:8: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "prefix-suffix-hashes:9: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "prefix-suffix-hashes:10: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "prefix-suffix-hashes:11: set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4
set probePhase to "prefix-suffix-hashes:12: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "prefix-suffix-hashes:13: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "prefix-suffix-hashes:14: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "prefix-suffix-hashes:15: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "prefix-suffix-hashes:16: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "prefix-suffix-hashes:17: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "prefix-suffix-hashes:18: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "prefix-suffix-hashes:19: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "prefix-suffix-hashes:20: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "prefix-suffix-hashes:21: set qualityPrefixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityPrefixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "prefix-suffix-hashes:22: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "prefix-suffix-hashes:23: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "prefix-suffix-hashes:24: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "prefix-suffix-hashes:25: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "prefix-suffix-hashes:26: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "prefix-suffix-hashes:27: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "prefix-suffix-hashes:28: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "prefix-suffix-hashes:29: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "prefix-suffix-hashes:30: set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4
set probePhase to "prefix-suffix-hashes:31: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "prefix-suffix-hashes:32: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "prefix-suffix-hashes:33: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "prefix-suffix-hashes:34: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "prefix-suffix-hashes:35: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "prefix-suffix-hashes:36: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "prefix-suffix-hashes:37: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "prefix-suffix-hashes:38: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "prefix-suffix-hashes:39: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "prefix-suffix-hashes:40: set qualitySuffixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualitySuffixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "prefix-suffix-hashes:41: set qualityState to {{\"quality-state\",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}"
log probePhase
set qualityState to {{"quality-state",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
