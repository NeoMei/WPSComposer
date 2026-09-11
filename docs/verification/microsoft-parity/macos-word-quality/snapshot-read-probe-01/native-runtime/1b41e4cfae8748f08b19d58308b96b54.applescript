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
set probePhase to "0: set qualityOldOrdinal to 0"
set qualityOldOrdinal to 0
repeat with qi from 1 to count tables of boundDoc
set probePhase to "2: set qt to table qi of boundDoc"
set qt to table qi of boundDoc
set probePhase to "3: set qa to start of content of text object of qt"
set qa to start of content of text object of qt
set probePhase to "4: set qz to end of content of text object of qt"
set qz to end of content of text object of qt
set probePhase to "5: if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then"
if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then
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
set probePhase to "14: set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4
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
set probePhase to "24: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "25: if qa >= qualityPoint + qualityDelta then"
if qa >= qualityPoint + qualityDelta then
set probePhase to "26: set qa to qa - qualityDelta"
set qa to qa - qualityDelta
set probePhase to "27: set qz to qz - qualityDelta"
set qz to qz - qualityDelta
end if
set probePhase to "29: set qualityOldOrdinal to qualityOldOrdinal + 1"
set qualityOldOrdinal to qualityOldOrdinal + 1
set probePhase to "30: set qualityRowFlags to {}"
set qualityRowFlags to {}
repeat with qrow in rows of qt
set probePhase to "32: set end of qualityRowFlags to allow break across pages of qrow"
set end of qualityRowFlags to allow break across pages of qrow
end repeat
set probePhase to "34: set end of qualityState to {\"table\",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}"
set end of qualityState to {"table",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}
end if
end repeat
set nativeRows to {{"read-ok",qualityState}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
