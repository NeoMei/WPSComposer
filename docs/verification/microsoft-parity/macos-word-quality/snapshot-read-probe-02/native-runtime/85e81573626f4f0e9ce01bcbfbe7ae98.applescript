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
set probePhase to "complete-snapshot-candidate:0: set qe to end of content of text object of boundDoc"
log probePhase
set qe to end of content of text object of boundDoc
set probePhase to "complete-snapshot-candidate:1: set qp to create range boundDoc start 0 end qualityPoint"
log probePhase
set qp to create range boundDoc start 0 end qualityPoint
set probePhase to "complete-snapshot-candidate:2: set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe"
log probePhase
set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe
set probePhase to "complete-snapshot-candidate:3: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:4: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:5: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:6: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:7: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:8: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:9: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:10: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:11: set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:12: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:13: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:14: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:15: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:16: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:17: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:18: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:19: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:20: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:21: set qualityPrefixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityPrefixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:22: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:23: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:24: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:25: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:26: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:27: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:28: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:29: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:30: set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:31: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:32: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:33: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:34: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:35: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:36: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:37: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:38: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:39: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:40: set qualitySuffixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualitySuffixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:41: set qualityState to {{\"quality-state\",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}"
log probePhase
set qualityState to {{"quality-state",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}
set probePhase to "complete-snapshot-candidate:42: set qualityLayout to {}"
log probePhase
set qualityLayout to {}
set probePhase to "complete-snapshot-candidate:43: repeat with qi from 1 to count sections of boundDoc"
log probePhase
repeat with qi from 1 to count sections of boundDoc
set probePhase to "complete-snapshot-candidate:44: set qsection to section qi of boundDoc"
log probePhase
set qsection to section qi of boundDoc
set probePhase to "complete-snapshot-candidate:45: set qsetup to page setup of qsection"
log probePhase
set qsetup to page setup of qsection
set probePhase to "complete-snapshot-candidate:46: set end of qualityLayout to {\"section\",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of "
log probePhase
set end of qualityLayout to {"section",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,header distance of qsetup,footer distance of qsetup,gutter of qsetup,count text columns of qsetup}
set probePhase to "complete-snapshot-candidate:47: repeat with qindex in {header footer primary,header footer first page,header footer even pages}"
log probePhase
repeat with qindex in {header footer primary,header footer first page,header footer even pages}
set probePhase to "complete-snapshot-candidate:48: set qheader to get header qsection index qindex"
log probePhase
set qheader to get header qsection index qindex
set probePhase to "complete-snapshot-candidate:49: set qfooter to get footer qsection index qindex"
log probePhase
set qfooter to get footer qsection index qindex
set probePhase to "complete-snapshot-candidate:50: repeat with qpart in {qheader,qfooter}"
log probePhase
repeat with qpart in {qheader,qfooter}
set probePhase to "complete-snapshot-candidate:51: if (count shapes of qpart) is not 0 then error \"WPSC_QUALITY_DRAWING_UNVERIFIED\""
log probePhase
if (count shapes of qpart) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"
set probePhase to "complete-snapshot-candidate:52: set qr to text object of qpart"
log probePhase
set qr to text object of qpart
set probePhase to "complete-snapshot-candidate:53: set end of qualityLayout to {\"page-part\",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as "
log probePhase
set end of qualityLayout to {"page-part",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as text}
set probePhase to "complete-snapshot-candidate:54: repeat with qfield in fields of qr"
log probePhase
repeat with qfield in fields of qr
set probePhase to "complete-snapshot-candidate:55: set end of qualityLayout to {\"page-field\",field type of qfield as text,content of field code of qfield as text,content of result range of qf"
log probePhase
set end of qualityLayout to {"page-field",field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield,end of content of field code of qfield,start of content of result range of qfield,end of content of result range of qfield,locked of qfield}
end repeat
end repeat
end repeat
end repeat
set probePhase to "complete-snapshot-candidate:60: repeat with qi from 1 to count list templates of boundDoc"
log probePhase
repeat with qi from 1 to count list templates of boundDoc
set probePhase to "complete-snapshot-candidate:61: set qtemplate to list template qi of boundDoc"
log probePhase
set qtemplate to list template qi of boundDoc
set probePhase to "complete-snapshot-candidate:62: set end of qualityLayout to {\"list-template\",qi as integer,name of qtemplate as text,outline numbered of qtemplate}"
log probePhase
set end of qualityLayout to {"list-template",qi as integer,name of qtemplate as text,outline numbered of qtemplate}
set probePhase to "complete-snapshot-candidate:63: repeat with qlevel in list levels of qtemplate"
log probePhase
repeat with qlevel in list levels of qtemplate
set probePhase to "complete-snapshot-candidate:64: set end of qualityLayout to {\"list-level\",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style "
log probePhase
set end of qualityLayout to {"list-level",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style of qlevel as text,start at of qlevel,reset on higher of qlevel,number position of qlevel,text position of qlevel,tab position of qlevel,trailing character of qlevel as text,list level alignment of qlevel as text}
end repeat
end repeat
set probePhase to "complete-snapshot-candidate:67: set qualityStyleNames to get name local of every Word style of boundDoc"
log probePhase
set qualityStyleNames to get name local of every Word style of boundDoc
set probePhase to "complete-snapshot-candidate:68: repeat with qualityStyleOrdinal from 1 to count qualityStyleNames"
log probePhase
repeat with qualityStyleOrdinal from 1 to count qualityStyleNames
set probePhase to "complete-snapshot-candidate:69: set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text"
log probePhase
set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text
set probePhase to "complete-snapshot-candidate:70: set qstyle to Word style qualityStyleName of boundDoc"
log probePhase
set qstyle to Word style qualityStyleName of boundDoc
set probePhase to "complete-snapshot-candidate:71: set end of qualityLayout to {\"style\",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}"
log probePhase
set end of qualityLayout to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}
end repeat
set probePhase to "complete-snapshot-candidate:73: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:74: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:75: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:76: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:77: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:78: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:79: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:80: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:81: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:82: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:83: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:84: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:85: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:86: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:87: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:88: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:89: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:90: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:91: set qualityLayoutHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityLayoutHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:92: set end of qualityState to {\"layout\",qualityLayoutHash}"
log probePhase
set end of qualityState to {"layout",qualityLayoutHash}
set probePhase to "complete-snapshot-candidate:93: repeat with qi from 1 to count paragraphs of boundDoc"
log probePhase
repeat with qi from 1 to count paragraphs of boundDoc
set probePhase to "complete-snapshot-candidate:94: set qr to text object of paragraph qi of boundDoc"
log probePhase
set qr to text object of paragraph qi of boundDoc
set probePhase to "complete-snapshot-candidate:95: set qa to start of content of qr"
log probePhase
set qa to start of content of qr
set probePhase to "complete-snapshot-candidate:96: set qz to end of content of qr"
log probePhase
set qz to end of content of qr
set probePhase to "complete-snapshot-candidate:97: if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then"
log probePhase
if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-candidate:98: if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then"
log probePhase
if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-candidate:99: set qa to qualityPoint + qualityDelta"
log probePhase
set qa to qualityPoint + qualityDelta
set probePhase to "complete-snapshot-candidate:100: set qr to create range boundDoc start qa end qz"
log probePhase
set qr to create range boundDoc start qa end qz
end if
set probePhase to "complete-snapshot-candidate:102: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:103: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:104: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:105: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:106: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:107: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:108: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:109: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:110: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:111: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:112: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:113: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:114: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:115: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:116: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:117: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:118: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:119: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:120: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:121: set qpf to paragraph format of qr"
log probePhase
set qpf to paragraph format of qr
set probePhase to "complete-snapshot-candidate:122: set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format righ"
log probePhase
set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,space before of qpf,space after of qpf,line spacing of qpf,line spacing rule of qpf as text,alignment of qpf as text,keep with next of qpf,keep together of qpf,widow control of qpf,outline level of qpf as text,character unit first line indent of qpf,list type of list format of qr as text,list level number of list format of qr,list value of list format of qr,list string of list format of qr as text}}
set probePhase to "complete-snapshot-candidate:123: repeat with qc from qa to qz - 1"
log probePhase
repeat with qc from qa to qz - 1
set probePhase to "complete-snapshot-candidate:124: set qcr to create range boundDoc start qc end (qc + 1)"
log probePhase
set qcr to create range boundDoc start qc end (qc + 1)
set probePhase to "complete-snapshot-candidate:125: set qcf to font object of qcr"
log probePhase
set qcf to font object of qcr
set probePhase to "complete-snapshot-candidate:126: set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,backgroun"
log probePhase
set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of shading of qcr}
end repeat
set probePhase to "complete-snapshot-candidate:128: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:129: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:130: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:131: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:132: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:133: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:134: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:135: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:136: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:137: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:138: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:139: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:140: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:141: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:142: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:143: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:144: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:145: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:146: set qualityFormatHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityFormatHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:147: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
log probePhase
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "complete-snapshot-candidate:148: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
log probePhase
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "complete-snapshot-candidate:149: set end of qualityState to {\"paragraph\",qa,qz,qualityTextHash,qualityFormatHash}"
log probePhase
set end of qualityState to {"paragraph",qa,qz,qualityTextHash,qualityFormatHash}
end if
end repeat
set probePhase to "complete-snapshot-candidate:152: repeat with qi from 1 to count fields of boundDoc"
log probePhase
repeat with qi from 1 to count fields of boundDoc
set probePhase to "complete-snapshot-candidate:153: set qf to field qi of boundDoc"
log probePhase
set qf to field qi of boundDoc
set probePhase to "complete-snapshot-candidate:154: set qa to start of content of field code of qf"
log probePhase
set qa to start of content of field code of qf
set probePhase to "complete-snapshot-candidate:155: set qz to end of content of field code of qf"
log probePhase
set qz to end of content of field code of qf
set probePhase to "complete-snapshot-candidate:156: set qra to start of content of result range of qf"
log probePhase
set qra to start of content of result range of qf
set probePhase to "complete-snapshot-candidate:157: set qrz to end of content of result range of qf"
log probePhase
set qrz to end of content of result range of qf
set probePhase to "complete-snapshot-candidate:158: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:159: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:160: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:161: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:162: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:163: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:164: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:165: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:166: set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:167: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:168: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:169: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:170: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:171: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:172: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:173: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:174: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:175: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:176: set qualityCodeHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityCodeHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:177: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:178: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:179: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:180: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:181: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:182: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:183: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:184: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:185: set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:186: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:187: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:188: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:189: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:190: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:191: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:192: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:193: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:194: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:195: set qualityResultHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityResultHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:196: if qa >= qualityPoint + qualityDelta then"
log probePhase
if qa >= qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-candidate:197: set qa to qa - qualityDelta"
log probePhase
set qa to qa - qualityDelta
set probePhase to "complete-snapshot-candidate:198: set qz to qz - qualityDelta"
log probePhase
set qz to qz - qualityDelta
set probePhase to "complete-snapshot-candidate:199: set qra to qra - qualityDelta"
log probePhase
set qra to qra - qualityDelta
set probePhase to "complete-snapshot-candidate:200: set qrz to qrz - qualityDelta"
log probePhase
set qrz to qrz - qualityDelta
end if
set probePhase to "complete-snapshot-candidate:202: set end of qualityState to {\"field\",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}"
log probePhase
set end of qualityState to {"field",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}
end repeat
set probePhase to "complete-snapshot-candidate:204: set qualityOldOrdinal to 0"
log probePhase
set qualityOldOrdinal to 0
set probePhase to "complete-snapshot-candidate:205: repeat with qi from 1 to count tables of boundDoc"
log probePhase
repeat with qi from 1 to count tables of boundDoc
set probePhase to "complete-snapshot-candidate:206: set qt to table qi of boundDoc"
log probePhase
set qt to table qi of boundDoc
set probePhase to "complete-snapshot-candidate:207: set qa to start of content of text object of qt"
log probePhase
set qa to start of content of text object of qt
set probePhase to "complete-snapshot-candidate:208: set qz to end of content of text object of qt"
log probePhase
set qz to end of content of text object of qt
set probePhase to "complete-snapshot-candidate:209: if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then"
log probePhase
if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-candidate:210: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:211: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:212: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:213: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:214: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:215: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:216: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:217: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:218: set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:219: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:220: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:221: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:222: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:223: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:224: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:225: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:226: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:227: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:228: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:229: if qa >= qualityPoint + qualityDelta then"
log probePhase
if qa >= qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-candidate:230: set qa to qa - qualityDelta"
log probePhase
set qa to qa - qualityDelta
set probePhase to "complete-snapshot-candidate:231: set qz to qz - qualityDelta"
log probePhase
set qz to qz - qualityDelta
end if
set probePhase to "complete-snapshot-candidate:233: set qualityOldOrdinal to qualityOldOrdinal + 1"
log probePhase
set qualityOldOrdinal to qualityOldOrdinal + 1
set probePhase to "complete-snapshot-candidate:234: set qualityRowFlags to {}"
log probePhase
set qualityRowFlags to {}
set probePhase to "complete-snapshot-candidate:235: repeat with qualityRowOrdinal from 1 to count rows of qt"
log probePhase
repeat with qualityRowOrdinal from 1 to count rows of qt
set probePhase to "complete-snapshot-candidate:236: set qrow to row qualityRowOrdinal of qt"
log probePhase
set qrow to row qualityRowOrdinal of qt
set probePhase to "complete-snapshot-candidate:237: set end of qualityRowFlags to allow break across pages of qrow"
log probePhase
set end of qualityRowFlags to allow break across pages of qrow
end repeat
set probePhase to "complete-snapshot-candidate:239: set end of qualityState to {\"table\",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}"
log probePhase
set end of qualityState to {"table",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}
end if
end repeat
set probePhase to "complete-snapshot-candidate:242: repeat with qi from 1 to count bookmarks of boundDoc"
log probePhase
repeat with qi from 1 to count bookmarks of boundDoc
set probePhase to "complete-snapshot-candidate:243: set qb to bookmark qi of boundDoc"
log probePhase
set qb to bookmark qi of boundDoc
set probePhase to "complete-snapshot-candidate:244: set qa to start of bookmark of qb"
log probePhase
set qa to start of bookmark of qb
set probePhase to "complete-snapshot-candidate:245: set qz to end of bookmark of qb"
log probePhase
set qz to end of bookmark of qb
set probePhase to "complete-snapshot-candidate:246: set qhashStart to qa"
log probePhase
set qhashStart to qa
set probePhase to "complete-snapshot-candidate:247: if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta"
log probePhase
if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta
set probePhase to "complete-snapshot-candidate:248: set qr to create range boundDoc start qhashStart end qz"
log probePhase
set qr to create range boundDoc start qhashStart end qz
set probePhase to "complete-snapshot-candidate:249: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-candidate:250: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-candidate:251: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-candidate:252: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:253: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-candidate:254: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-candidate:255: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-candidate:256: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-candidate:257: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-candidate:258: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-candidate:259: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-candidate:260: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-candidate:261: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-candidate:262: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-candidate:263: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:264: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-candidate:265: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:266: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-candidate:267: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-candidate:268: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
log probePhase
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "complete-snapshot-candidate:269: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
log probePhase
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "complete-snapshot-candidate:270: set end of qualityState to {\"bookmark\",name of qb as text,qa,qz,qualityTextHash}"
log probePhase
set end of qualityState to {"bookmark",name of qb as text,qa,qz,qualityTextHash}
end repeat
set probePhase to "complete-snapshot-candidate:272: set end of qualityState to {\"quality-state-end\"}"
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
