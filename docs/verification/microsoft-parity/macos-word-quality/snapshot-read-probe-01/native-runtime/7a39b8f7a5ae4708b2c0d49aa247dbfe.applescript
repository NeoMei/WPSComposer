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
repeat with qi from 1 to count paragraphs of boundDoc
set probePhase to "1: set qr to text object of paragraph qi of boundDoc"
set qr to text object of paragraph qi of boundDoc
set probePhase to "2: set qa to start of content of qr"
set qa to start of content of qr
set probePhase to "3: set qz to end of content of qr"
set qz to end of content of qr
set probePhase to "4: if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then"
if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then
set probePhase to "5: if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then"
if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then
set probePhase to "6: set qa to qualityPoint + qualityDelta"
set qa to qualityPoint + qualityDelta
set probePhase to "7: set qr to create range boundDoc start qa end qz"
set qr to create range boundDoc start qa end qz
end if
set probePhase to "9: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "10: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "11: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "12: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "13: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "14: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "15: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "16: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "17: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "18: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "19: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "20: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "21: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "22: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "23: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "24: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "25: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "26: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "27: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "28: set qpf to paragraph format of qr"
set qpf to paragraph format of qr
set probePhase to "29: set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,space before of qpf,space after of qpf,line spacing of qpf,line "
set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,space before of qpf,space after of qpf,line spacing of qpf,line spacing rule of qpf as text,alignment of qpf as text,keep with next of qpf,keep together of qpf,widow control of qpf,outline level of qpf as text,character unit first line indent of qpf,list type of list format of qr as text,list level number of list format of qr,list value of list format of qr,list string of list format of qr as text}}
repeat with qc from qa to qz - 1
set probePhase to "31: set qcr to create range boundDoc start qc end (qc + 1)"
set qcr to create range boundDoc start qc end (qc + 1)
set probePhase to "32: set qcf to font object of qcr"
set qcf to font object of qcr
set probePhase to "33: set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of shading of qcr}"
set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of shading of qcr}
end repeat
set probePhase to "35: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "36: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "37: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "38: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "39: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "40: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "41: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "42: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "43: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4
set probePhase to "44: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "45: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "46: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "47: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "48: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "49: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "50: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "51: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "52: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "53: set qualityFormatHash to text 1 thru 64 of recoveryDigestText"
set qualityFormatHash to text 1 thru 64 of recoveryDigestText
set probePhase to "54: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "55: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "56: set end of qualityState to {\"paragraph\",qa,qz,qualityTextHash,qualityFormatHash}"
set end of qualityState to {"paragraph",qa,qz,qualityTextHash,qualityFormatHash}
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
