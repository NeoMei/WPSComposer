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
repeat with qi from 1 to count fields of boundDoc
set probePhase to "1: set qf to field qi of boundDoc"
set qf to field qi of boundDoc
set probePhase to "2: set qa to start of content of field code of qf"
set qa to start of content of field code of qf
set probePhase to "3: set qz to end of content of field code of qf"
set qz to end of content of field code of qf
set probePhase to "4: set qra to start of content of result range of qf"
set qra to start of content of result range of qf
set probePhase to "5: set qrz to end of content of result range of qf"
set qrz to end of content of result range of qf
set probePhase to "6: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "7: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "8: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "9: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "10: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "11: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "12: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "13: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "14: set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4
set probePhase to "15: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "16: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "17: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "18: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "19: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "20: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "21: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "22: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "23: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "24: set qualityCodeHash to text 1 thru 64 of recoveryDigestText"
set qualityCodeHash to text 1 thru 64 of recoveryDigestText
set probePhase to "25: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "26: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "27: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "28: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "29: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "30: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "31: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "32: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "33: set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4
set probePhase to "34: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "35: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "36: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "37: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "38: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "39: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "40: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "41: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "42: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "43: set qualityResultHash to text 1 thru 64 of recoveryDigestText"
set qualityResultHash to text 1 thru 64 of recoveryDigestText
set probePhase to "44: if qa >= qualityPoint + qualityDelta then"
if qa >= qualityPoint + qualityDelta then
set probePhase to "45: set qa to qa - qualityDelta"
set qa to qa - qualityDelta
set probePhase to "46: set qz to qz - qualityDelta"
set qz to qz - qualityDelta
set probePhase to "47: set qra to qra - qualityDelta"
set qra to qra - qualityDelta
set probePhase to "48: set qrz to qrz - qualityDelta"
set qrz to qrz - qualityDelta
end if
set probePhase to "50: set end of qualityState to {\"field\",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}"
set end of qualityState to {"field",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}
end repeat
set nativeRows to {{"read-ok",qualityState}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
