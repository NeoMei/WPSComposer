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
set probePhase to "0: set qe to end of content of text object of boundDoc"
set qe to end of content of text object of boundDoc
set probePhase to "1: set qp to create range boundDoc start 0 end qualityPoint"
set qp to create range boundDoc start 0 end qualityPoint
set probePhase to "2: set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe"
set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe
set probePhase to "3: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "4: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "5: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "6: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "7: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "8: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "9: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "10: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "11: set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4
set probePhase to "12: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "13: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "14: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "15: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "16: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "17: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "18: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "19: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "20: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "21: set qualityPrefixHash to text 1 thru 64 of recoveryDigestText"
set qualityPrefixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "22: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "23: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "24: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "25: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "26: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "27: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "28: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "29: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "30: set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4
set probePhase to "31: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "32: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "33: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "34: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "35: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "36: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "37: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "38: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "39: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "40: set qualitySuffixHash to text 1 thru 64 of recoveryDigestText"
set qualitySuffixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "41: set qualityState to {{\"quality-state\",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}"
set qualityState to {{"quality-state",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}
set probePhase to "42: set qualityLayout to {}"
set qualityLayout to {}
repeat with qi from 1 to count sections of boundDoc
set probePhase to "44: set qsection to section qi of boundDoc"
set qsection to section qi of boundDoc
set probePhase to "45: set qsetup to page setup of qsection"
set qsetup to page setup of qsection
set probePhase to "46: set end of qualityLayout to {\"section\",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,head"
set end of qualityLayout to {"section",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,header distance of qsetup,footer distance of qsetup,gutter of qsetup,count text columns of qsetup}
repeat with qindex in {header footer primary,header footer first page,header footer even pages}
set probePhase to "48: set qheader to get header qsection index qindex"
set qheader to get header qsection index qindex
set probePhase to "49: set qfooter to get footer qsection index qindex"
set qfooter to get footer qsection index qindex
repeat with qpart in {qheader,qfooter}
set probePhase to "51: if (count shapes of qpart) is not 0 then error \"WPSC_QUALITY_DRAWING_UNVERIFIED\""
if (count shapes of qpart) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"
set probePhase to "52: set qr to text object of qpart"
set qr to text object of qpart
set probePhase to "53: set end of qualityLayout to {\"page-part\",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as text}"
set end of qualityLayout to {"page-part",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as text}
repeat with qfield in fields of qr
set probePhase to "55: set end of qualityLayout to {\"page-field\",field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield,end of content of field co"
set end of qualityLayout to {"page-field",field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield,end of content of field code of qfield,start of content of result range of qfield,end of content of result range of qfield,locked of qfield}
end repeat
end repeat
end repeat
end repeat
repeat with qi from 1 to count list templates of boundDoc
set probePhase to "61: set qtemplate to list template qi of boundDoc"
set qtemplate to list template qi of boundDoc
set probePhase to "62: set end of qualityLayout to {\"list-template\",qi as integer,name of qtemplate as text,outline numbered of qtemplate}"
set end of qualityLayout to {"list-template",qi as integer,name of qtemplate as text,outline numbered of qtemplate}
repeat with qlevel in list levels of qtemplate
set probePhase to "64: set end of qualityLayout to {\"list-level\",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style of qlevel as text,start at of qlevel,reset on higher of qlevel,number position o"
set end of qualityLayout to {"list-level",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style of qlevel as text,start at of qlevel,reset on higher of qlevel,number position of qlevel,text position of qlevel,tab position of qlevel,trailing character of qlevel as text,list level alignment of qlevel as text}
end repeat
end repeat
repeat with qstyle in Word styles of boundDoc
set probePhase to "68: set end of qualityLayout to {\"style\",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}"
set end of qualityLayout to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}
end repeat
set probePhase to "70: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "71: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "72: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "73: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "74: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "75: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "76: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "77: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "78: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4
set probePhase to "79: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "80: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "81: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "82: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "83: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "84: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "85: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "86: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "87: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "88: set qualityLayoutHash to text 1 thru 64 of recoveryDigestText"
set qualityLayoutHash to text 1 thru 64 of recoveryDigestText
set probePhase to "89: set end of qualityState to {\"layout\",qualityLayoutHash}"
set end of qualityState to {"layout",qualityLayoutHash}
repeat with qi from 1 to count paragraphs of boundDoc
set probePhase to "91: set qr to text object of paragraph qi of boundDoc"
set qr to text object of paragraph qi of boundDoc
set probePhase to "92: set qa to start of content of qr"
set qa to start of content of qr
set probePhase to "93: set qz to end of content of qr"
set qz to end of content of qr
set probePhase to "94: if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then"
if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then
set probePhase to "95: if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then"
if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then
set probePhase to "96: set qa to qualityPoint + qualityDelta"
set qa to qualityPoint + qualityDelta
set probePhase to "97: set qr to create range boundDoc start qa end qz"
set qr to create range boundDoc start qa end qz
end if
set probePhase to "99: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "100: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "101: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "102: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "103: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "104: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "105: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "106: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "107: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "108: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "109: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "110: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "111: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "112: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "113: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "114: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "115: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "116: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "117: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "118: set qpf to paragraph format of qr"
set qpf to paragraph format of qr
set probePhase to "119: set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,space before of qpf,space after of qpf,line spacing of qpf,line "
set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,space before of qpf,space after of qpf,line spacing of qpf,line spacing rule of qpf as text,alignment of qpf as text,keep with next of qpf,keep together of qpf,widow control of qpf,outline level of qpf as text,character unit first line indent of qpf,list type of list format of qr as text,list level number of list format of qr,list value of list format of qr,list string of list format of qr as text}}
repeat with qc from qa to qz - 1
set probePhase to "121: set qcr to create range boundDoc start qc end (qc + 1)"
set qcr to create range boundDoc start qc end (qc + 1)
set probePhase to "122: set qcf to font object of qcr"
set qcf to font object of qcr
set probePhase to "123: set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of shading of qcr}"
set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of shading of qcr}
end repeat
set probePhase to "125: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "126: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "127: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "128: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "129: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "130: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "131: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "132: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "133: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4
set probePhase to "134: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "135: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "136: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "137: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "138: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "139: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "140: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "141: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "142: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "143: set qualityFormatHash to text 1 thru 64 of recoveryDigestText"
set qualityFormatHash to text 1 thru 64 of recoveryDigestText
set probePhase to "144: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "145: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "146: set end of qualityState to {\"paragraph\",qa,qz,qualityTextHash,qualityFormatHash}"
set end of qualityState to {"paragraph",qa,qz,qualityTextHash,qualityFormatHash}
end if
end repeat
repeat with qi from 1 to count fields of boundDoc
set probePhase to "150: set qf to field qi of boundDoc"
set qf to field qi of boundDoc
set probePhase to "151: set qa to start of content of field code of qf"
set qa to start of content of field code of qf
set probePhase to "152: set qz to end of content of field code of qf"
set qz to end of content of field code of qf
set probePhase to "153: set qra to start of content of result range of qf"
set qra to start of content of result range of qf
set probePhase to "154: set qrz to end of content of result range of qf"
set qrz to end of content of result range of qf
set probePhase to "155: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "156: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "157: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "158: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "159: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "160: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "161: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "162: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "163: set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4
set probePhase to "164: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "165: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "166: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "167: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "168: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "169: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "170: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "171: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "172: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "173: set qualityCodeHash to text 1 thru 64 of recoveryDigestText"
set qualityCodeHash to text 1 thru 64 of recoveryDigestText
set probePhase to "174: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "175: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "176: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "177: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "178: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "179: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "180: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "181: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "182: set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4
set probePhase to "183: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "184: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "185: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "186: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "187: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "188: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "189: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "190: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "191: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "192: set qualityResultHash to text 1 thru 64 of recoveryDigestText"
set qualityResultHash to text 1 thru 64 of recoveryDigestText
set probePhase to "193: if qa >= qualityPoint + qualityDelta then"
if qa >= qualityPoint + qualityDelta then
set probePhase to "194: set qa to qa - qualityDelta"
set qa to qa - qualityDelta
set probePhase to "195: set qz to qz - qualityDelta"
set qz to qz - qualityDelta
set probePhase to "196: set qra to qra - qualityDelta"
set qra to qra - qualityDelta
set probePhase to "197: set qrz to qrz - qualityDelta"
set qrz to qrz - qualityDelta
end if
set probePhase to "199: set end of qualityState to {\"field\",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}"
set end of qualityState to {"field",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}
end repeat
set probePhase to "201: set qualityOldOrdinal to 0"
set qualityOldOrdinal to 0
repeat with qi from 1 to count tables of boundDoc
set probePhase to "203: set qt to table qi of boundDoc"
set qt to table qi of boundDoc
set probePhase to "204: set qa to start of content of text object of qt"
set qa to start of content of text object of qt
set probePhase to "205: set qz to end of content of text object of qt"
set qz to end of content of text object of qt
set probePhase to "206: if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then"
if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then
set probePhase to "207: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "208: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "209: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "210: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "211: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "212: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "213: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "214: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "215: set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4
set probePhase to "216: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "217: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "218: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "219: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "220: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "221: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "222: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "223: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "224: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "225: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "226: if qa >= qualityPoint + qualityDelta then"
if qa >= qualityPoint + qualityDelta then
set probePhase to "227: set qa to qa - qualityDelta"
set qa to qa - qualityDelta
set probePhase to "228: set qz to qz - qualityDelta"
set qz to qz - qualityDelta
end if
set probePhase to "230: set qualityOldOrdinal to qualityOldOrdinal + 1"
set qualityOldOrdinal to qualityOldOrdinal + 1
set probePhase to "231: set qualityRowFlags to {}"
set qualityRowFlags to {}
repeat with qrow in rows of qt
set probePhase to "233: set end of qualityRowFlags to allow break across pages of qrow"
set end of qualityRowFlags to allow break across pages of qrow
end repeat
set probePhase to "235: set end of qualityState to {\"table\",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}"
set end of qualityState to {"table",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}
end if
end repeat
repeat with qi from 1 to count bookmarks of boundDoc
set probePhase to "239: set qb to bookmark qi of boundDoc"
set qb to bookmark qi of boundDoc
set probePhase to "240: set qa to start of bookmark of qb"
set qa to start of bookmark of qb
set probePhase to "241: set qz to end of bookmark of qb"
set qz to end of bookmark of qb
set probePhase to "242: set qhashStart to qa"
set qhashStart to qa
set probePhase to "243: if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta"
if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta
set probePhase to "244: set qr to create range boundDoc start qhashStart end qz"
set qr to create range boundDoc start qhashStart end qz
set probePhase to "245: set recoveryTask to current application's NSTask's alloc()'s init()"
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "246: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "247: recoveryTask's setArguments:{\"-a\", \"256\"}"
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "248: set recoveryInput to current application's NSPipe's pipe()"
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "249: set recoveryOutput to current application's NSPipe's pipe()"
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "250: recoveryTask's setStandardInput:recoveryInput"
recoveryTask's setStandardInput:recoveryInput
set probePhase to "251: recoveryTask's setStandardOutput:recoveryOutput"
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "252: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "253: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "254: recoveryTask's |launch|()"
recoveryTask's |launch|()
set probePhase to "255: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "256: (recoveryInput's fileHandleForWriting())'s closeFile()"
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "257: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "258: recoveryTask's waitUntilExit()"
recoveryTask's waitUntilExit()
set probePhase to "259: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "260: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "261: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "262: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "263: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "264: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "265: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "266: set end of qualityState to {\"bookmark\",name of qb as text,qa,qz,qualityTextHash}"
set end of qualityState to {"bookmark",name of qb as text,qa,qz,qualityTextHash}
end repeat
set probePhase to "268: set end of qualityState to {\"quality-state-end\"}"
set end of qualityState to {"quality-state-end"}
set nativeRows to {{"read-ok",qualityState}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
