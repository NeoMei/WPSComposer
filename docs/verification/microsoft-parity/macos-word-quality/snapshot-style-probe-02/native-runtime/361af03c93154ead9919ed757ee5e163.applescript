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
set boundDoc to document "document-85eb770d18014882a21f324998578555.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-u3x1svek/document-85eb770d18014882a21f324998578555.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "complete-snapshot-defined-styles:0: set qe to end of content of text object of boundDoc"
log probePhase
set qe to end of content of text object of boundDoc
set probePhase to "complete-snapshot-defined-styles:1: set qp to create range boundDoc start 0 end qualityPoint"
log probePhase
set qp to create range boundDoc start 0 end qualityPoint
set probePhase to "complete-snapshot-defined-styles:2: set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe"
log probePhase
set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe
set probePhase to "complete-snapshot-defined-styles:3: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:4: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:5: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:6: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:7: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:8: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:9: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:10: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:11: set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:12: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:13: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:14: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:15: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:16: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:17: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:18: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:19: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:20: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:21: set qualityPrefixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityPrefixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:22: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:23: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:24: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:25: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:26: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:27: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:28: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:29: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:30: set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:31: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:32: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:33: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:34: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:35: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:36: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:37: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:38: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:39: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:40: set qualitySuffixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualitySuffixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:41: set qualityState to {{\"quality-state\",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}"
log probePhase
set qualityState to {{"quality-state",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}
set probePhase to "complete-snapshot-defined-styles:42: set qualityLayout to {}"
log probePhase
set qualityLayout to {}
set probePhase to "complete-snapshot-defined-styles:43: repeat with qi from 1 to count sections of boundDoc"
log probePhase
repeat with qi from 1 to count sections of boundDoc
set probePhase to "complete-snapshot-defined-styles:44: set qsection to section qi of boundDoc"
log probePhase
set qsection to section qi of boundDoc
set probePhase to "complete-snapshot-defined-styles:45: set qsetup to page setup of qsection"
log probePhase
set qsetup to page setup of qsection
set probePhase to "complete-snapshot-defined-styles:46: set end of qualityLayout to {\"section\",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of "
log probePhase
set end of qualityLayout to {"section",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,header distance of qsetup,footer distance of qsetup,gutter of qsetup,count text columns of qsetup}
set probePhase to "complete-snapshot-defined-styles:47: repeat with qindex in {header footer primary,header footer first page,header footer even pages}"
log probePhase
repeat with qindex in {header footer primary,header footer first page,header footer even pages}
set probePhase to "complete-snapshot-defined-styles:48: set qheader to get header qsection index qindex"
log probePhase
set qheader to get header qsection index qindex
set probePhase to "complete-snapshot-defined-styles:49: set qfooter to get footer qsection index qindex"
log probePhase
set qfooter to get footer qsection index qindex
set probePhase to "complete-snapshot-defined-styles:50: repeat with qpart in {qheader,qfooter}"
log probePhase
repeat with qpart in {qheader,qfooter}
set probePhase to "complete-snapshot-defined-styles:51: if (count shapes of qpart) is not 0 then error \"WPSC_QUALITY_DRAWING_UNVERIFIED\""
log probePhase
if (count shapes of qpart) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"
set probePhase to "complete-snapshot-defined-styles:52: set qr to text object of qpart"
log probePhase
set qr to text object of qpart
set probePhase to "complete-snapshot-defined-styles:53: set end of qualityLayout to {\"page-part\",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as "
log probePhase
set end of qualityLayout to {"page-part",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as text}
set probePhase to "complete-snapshot-defined-styles:54: repeat with qfield in fields of qr"
log probePhase
repeat with qfield in fields of qr
set probePhase to "complete-snapshot-defined-styles:55: set end of qualityLayout to {\"page-field\",field type of qfield as text,content of field code of qfield as text,content of result range of qf"
log probePhase
set end of qualityLayout to {"page-field",field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield,end of content of field code of qfield,start of content of result range of qfield,end of content of result range of qfield,locked of qfield}
end repeat
end repeat
end repeat
end repeat
set probePhase to "complete-snapshot-defined-styles:60: repeat with qi from 1 to count list templates of boundDoc"
log probePhase
repeat with qi from 1 to count list templates of boundDoc
set probePhase to "complete-snapshot-defined-styles:61: set qtemplate to list template qi of boundDoc"
log probePhase
set qtemplate to list template qi of boundDoc
set probePhase to "complete-snapshot-defined-styles:62: set end of qualityLayout to {\"list-template\",qi as integer,name of qtemplate as text,outline numbered of qtemplate}"
log probePhase
set end of qualityLayout to {"list-template",qi as integer,name of qtemplate as text,outline numbered of qtemplate}
set probePhase to "complete-snapshot-defined-styles:63: repeat with qlevel in list levels of qtemplate"
log probePhase
repeat with qlevel in list levels of qtemplate
set probePhase to "complete-snapshot-defined-styles:64: set end of qualityLayout to {\"list-level\",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style "
log probePhase
set end of qualityLayout to {"list-level",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style of qlevel as text,start at of qlevel,reset on higher of qlevel,number position of qlevel,text position of qlevel,tab position of qlevel,trailing character of qlevel as text,list level alignment of qlevel as text}
end repeat
end repeat
set probePhase to "complete-snapshot-defined-styles:67: set qualityStyleNames to get name local of every Word style of boundDoc"
log probePhase
set qualityStyleNames to get name local of every Word style of boundDoc
set probePhase to "complete-snapshot-defined-styles:68: set qualityStyleInUse to get in use of every Word style of boundDoc"
log probePhase
set qualityStyleInUse to get in use of every Word style of boundDoc
set probePhase to "complete-snapshot-defined-styles:69: set qualityStyleBuiltIn to get built in of every Word style of boundDoc"
log probePhase
set qualityStyleBuiltIn to get built in of every Word style of boundDoc
set probePhase to "complete-snapshot-defined-styles:70: if (count qualityStyleNames) is not (count qualityStyleInUse) or (count qualityStyleNames) is not (count qualityStyleBuiltIn) then error \"WP"
log probePhase
if (count qualityStyleNames) is not (count qualityStyleInUse) or (count qualityStyleNames) is not (count qualityStyleBuiltIn) then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set probePhase to "complete-snapshot-defined-styles:71: repeat with qualityStyleOrdinal from 1 to count qualityStyleNames"
log probePhase
repeat with qualityStyleOrdinal from 1 to count qualityStyleNames
set probePhase to "complete-snapshot-defined-styles:72: set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text"
log probePhase
set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text
set probePhase to "complete-snapshot-defined-styles:73: set qualityStyleUsed to item qualityStyleOrdinal of qualityStyleInUse"
log probePhase
set qualityStyleUsed to item qualityStyleOrdinal of qualityStyleInUse
set probePhase to "complete-snapshot-defined-styles:74: set qualityStyleBuiltin to item qualityStyleOrdinal of qualityStyleBuiltIn"
log probePhase
set qualityStyleBuiltin to item qualityStyleOrdinal of qualityStyleBuiltIn
set probePhase to "complete-snapshot-defined-styles:75: set end of qualityLayout to {\"style-state\",qualityStyleName,qualityStyleUsed,qualityStyleBuiltin}"
log probePhase
set end of qualityLayout to {"style-state",qualityStyleName,qualityStyleUsed,qualityStyleBuiltin}
set probePhase to "complete-snapshot-defined-styles:76: if qualityStyleUsed or not qualityStyleBuiltin then"
log probePhase
if qualityStyleUsed or not qualityStyleBuiltin then
set probePhase to "complete-snapshot-defined-styles:77: set qstyle to Word style qualityStyleName of boundDoc"
log probePhase
set qstyle to Word style qualityStyleName of boundDoc
set probePhase to "complete-snapshot-defined-styles:78: if (name local of qstyle as text) is not qualityStyleName then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\""
log probePhase
if (name local of qstyle as text) is not qualityStyleName then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set probePhase to "complete-snapshot-defined-styles:79: set end of qualityLayout to {\"style\",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}"
log probePhase
set end of qualityLayout to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}
end if
end repeat
set probePhase to "complete-snapshot-defined-styles:82: if (get name local of every Word style of boundDoc) is not qualityStyleNames then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\""
log probePhase
if (get name local of every Word style of boundDoc) is not qualityStyleNames then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set probePhase to "complete-snapshot-defined-styles:83: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:84: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:85: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:86: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:87: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:88: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:89: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:90: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:91: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:92: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:93: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:94: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:95: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:96: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:97: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:98: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:99: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:100: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:101: set qualityLayoutHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityLayoutHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:102: set end of qualityState to {\"layout\",qualityLayoutHash}"
log probePhase
set end of qualityState to {"layout",qualityLayoutHash}
set probePhase to "complete-snapshot-defined-styles:103: repeat with qi from 1 to count paragraphs of boundDoc"
log probePhase
repeat with qi from 1 to count paragraphs of boundDoc
set probePhase to "complete-snapshot-defined-styles:104: set qr to text object of paragraph qi of boundDoc"
log probePhase
set qr to text object of paragraph qi of boundDoc
set probePhase to "complete-snapshot-defined-styles:105: set qa to start of content of qr"
log probePhase
set qa to start of content of qr
set probePhase to "complete-snapshot-defined-styles:106: set qz to end of content of qr"
log probePhase
set qz to end of content of qr
set probePhase to "complete-snapshot-defined-styles:107: if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then"
log probePhase
if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-defined-styles:108: if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then"
log probePhase
if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-defined-styles:109: set qa to qualityPoint + qualityDelta"
log probePhase
set qa to qualityPoint + qualityDelta
set probePhase to "complete-snapshot-defined-styles:110: set qr to create range boundDoc start qa end qz"
log probePhase
set qr to create range boundDoc start qa end qz
end if
set probePhase to "complete-snapshot-defined-styles:112: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:113: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:114: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:115: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:116: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:117: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:118: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:119: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:120: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:121: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:122: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:123: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:124: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:125: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:126: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:127: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:128: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:129: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:130: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:131: set qpf to paragraph format of qr"
log probePhase
set qpf to paragraph format of qr
set probePhase to "complete-snapshot-defined-styles:132: set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format righ"
log probePhase
set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,space before of qpf,space after of qpf,line spacing of qpf,line spacing rule of qpf as text,alignment of qpf as text,keep with next of qpf,keep together of qpf,widow control of qpf,outline level of qpf as text,character unit first line indent of qpf,list type of list format of qr as text,list level number of list format of qr,list value of list format of qr,list string of list format of qr as text}}
set probePhase to "complete-snapshot-defined-styles:133: repeat with qc from qa to qz - 1"
log probePhase
repeat with qc from qa to qz - 1
set probePhase to "complete-snapshot-defined-styles:134: set qcr to create range boundDoc start qc end (qc + 1)"
log probePhase
set qcr to create range boundDoc start qc end (qc + 1)
set probePhase to "complete-snapshot-defined-styles:135: set qcf to font object of qcr"
log probePhase
set qcf to font object of qcr
set probePhase to "complete-snapshot-defined-styles:136: set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,backgroun"
log probePhase
set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of shading of qcr}
end repeat
set probePhase to "complete-snapshot-defined-styles:138: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:139: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:140: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:141: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:142: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:143: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:144: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:145: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:146: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:147: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:148: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:149: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:150: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:151: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:152: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:153: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:154: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:155: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:156: set qualityFormatHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityFormatHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:157: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
log probePhase
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "complete-snapshot-defined-styles:158: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
log probePhase
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "complete-snapshot-defined-styles:159: set end of qualityState to {\"paragraph\",qa,qz,qualityTextHash,qualityFormatHash}"
log probePhase
set end of qualityState to {"paragraph",qa,qz,qualityTextHash,qualityFormatHash}
end if
end repeat
set probePhase to "complete-snapshot-defined-styles:162: repeat with qi from 1 to count fields of boundDoc"
log probePhase
repeat with qi from 1 to count fields of boundDoc
set probePhase to "complete-snapshot-defined-styles:163: set qf to field qi of boundDoc"
log probePhase
set qf to field qi of boundDoc
set probePhase to "complete-snapshot-defined-styles:164: set qa to start of content of field code of qf"
log probePhase
set qa to start of content of field code of qf
set probePhase to "complete-snapshot-defined-styles:165: set qz to end of content of field code of qf"
log probePhase
set qz to end of content of field code of qf
set probePhase to "complete-snapshot-defined-styles:166: set qra to start of content of result range of qf"
log probePhase
set qra to start of content of result range of qf
set probePhase to "complete-snapshot-defined-styles:167: set qrz to end of content of result range of qf"
log probePhase
set qrz to end of content of result range of qf
set probePhase to "complete-snapshot-defined-styles:168: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:169: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:170: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:171: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:172: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:173: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:174: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:175: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:176: set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:177: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:178: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:179: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:180: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:181: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:182: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:183: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:184: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:185: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:186: set qualityCodeHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityCodeHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:187: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:188: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:189: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:190: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:191: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:192: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:193: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:194: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:195: set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:196: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:197: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:198: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:199: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:200: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:201: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:202: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:203: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:204: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:205: set qualityResultHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityResultHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:206: if qa >= qualityPoint + qualityDelta then"
log probePhase
if qa >= qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-defined-styles:207: set qa to qa - qualityDelta"
log probePhase
set qa to qa - qualityDelta
set probePhase to "complete-snapshot-defined-styles:208: set qz to qz - qualityDelta"
log probePhase
set qz to qz - qualityDelta
set probePhase to "complete-snapshot-defined-styles:209: set qra to qra - qualityDelta"
log probePhase
set qra to qra - qualityDelta
set probePhase to "complete-snapshot-defined-styles:210: set qrz to qrz - qualityDelta"
log probePhase
set qrz to qrz - qualityDelta
end if
set probePhase to "complete-snapshot-defined-styles:212: set end of qualityState to {\"field\",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}"
log probePhase
set end of qualityState to {"field",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}
end repeat
set probePhase to "complete-snapshot-defined-styles:214: set qualityOldOrdinal to 0"
log probePhase
set qualityOldOrdinal to 0
set probePhase to "complete-snapshot-defined-styles:215: repeat with qi from 1 to count tables of boundDoc"
log probePhase
repeat with qi from 1 to count tables of boundDoc
set probePhase to "complete-snapshot-defined-styles:216: set qt to table qi of boundDoc"
log probePhase
set qt to table qi of boundDoc
set probePhase to "complete-snapshot-defined-styles:217: set qa to start of content of text object of qt"
log probePhase
set qa to start of content of text object of qt
set probePhase to "complete-snapshot-defined-styles:218: set qz to end of content of text object of qt"
log probePhase
set qz to end of content of text object of qt
set probePhase to "complete-snapshot-defined-styles:219: if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then"
log probePhase
if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-defined-styles:220: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:221: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:222: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:223: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:224: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:225: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:226: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:227: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:228: set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:229: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:230: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:231: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:232: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:233: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:234: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:235: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:236: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:237: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:238: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:239: if qa >= qualityPoint + qualityDelta then"
log probePhase
if qa >= qualityPoint + qualityDelta then
set probePhase to "complete-snapshot-defined-styles:240: set qa to qa - qualityDelta"
log probePhase
set qa to qa - qualityDelta
set probePhase to "complete-snapshot-defined-styles:241: set qz to qz - qualityDelta"
log probePhase
set qz to qz - qualityDelta
end if
set probePhase to "complete-snapshot-defined-styles:243: set qualityOldOrdinal to qualityOldOrdinal + 1"
log probePhase
set qualityOldOrdinal to qualityOldOrdinal + 1
set probePhase to "complete-snapshot-defined-styles:244: set qualityRowFlags to {}"
log probePhase
set qualityRowFlags to {}
set probePhase to "complete-snapshot-defined-styles:245: repeat with qualityRowOrdinal from 1 to count rows of qt"
log probePhase
repeat with qualityRowOrdinal from 1 to count rows of qt
set probePhase to "complete-snapshot-defined-styles:246: set qrow to row qualityRowOrdinal of qt"
log probePhase
set qrow to row qualityRowOrdinal of qt
set probePhase to "complete-snapshot-defined-styles:247: set end of qualityRowFlags to allow break across pages of qrow"
log probePhase
set end of qualityRowFlags to allow break across pages of qrow
end repeat
set probePhase to "complete-snapshot-defined-styles:249: set end of qualityState to {\"table\",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}"
log probePhase
set end of qualityState to {"table",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}
end if
end repeat
set probePhase to "complete-snapshot-defined-styles:252: repeat with qi from 1 to count bookmarks of boundDoc"
log probePhase
repeat with qi from 1 to count bookmarks of boundDoc
set probePhase to "complete-snapshot-defined-styles:253: set qb to bookmark qi of boundDoc"
log probePhase
set qb to bookmark qi of boundDoc
set probePhase to "complete-snapshot-defined-styles:254: set qa to start of bookmark of qb"
log probePhase
set qa to start of bookmark of qb
set probePhase to "complete-snapshot-defined-styles:255: set qz to end of bookmark of qb"
log probePhase
set qz to end of bookmark of qb
set probePhase to "complete-snapshot-defined-styles:256: set qhashStart to qa"
log probePhase
set qhashStart to qa
set probePhase to "complete-snapshot-defined-styles:257: if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta"
log probePhase
if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta
set probePhase to "complete-snapshot-defined-styles:258: set qr to create range boundDoc start qhashStart end qz"
log probePhase
set qr to create range boundDoc start qhashStart end qz
set probePhase to "complete-snapshot-defined-styles:259: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "complete-snapshot-defined-styles:260: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "complete-snapshot-defined-styles:261: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "complete-snapshot-defined-styles:262: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:263: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "complete-snapshot-defined-styles:264: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "complete-snapshot-defined-styles:265: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "complete-snapshot-defined-styles:266: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "complete-snapshot-defined-styles:267: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "complete-snapshot-defined-styles:268: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "complete-snapshot-defined-styles:269: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "complete-snapshot-defined-styles:270: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "complete-snapshot-defined-styles:271: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "complete-snapshot-defined-styles:272: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "complete-snapshot-defined-styles:273: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:274: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "complete-snapshot-defined-styles:275: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:276: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "complete-snapshot-defined-styles:277: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "complete-snapshot-defined-styles:278: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
log probePhase
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "complete-snapshot-defined-styles:279: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
log probePhase
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "complete-snapshot-defined-styles:280: set end of qualityState to {\"bookmark\",name of qb as text,qa,qz,qualityTextHash}"
log probePhase
set end of qualityState to {"bookmark",name of qb as text,qa,qz,qualityTextHash}
end repeat
set probePhase to "complete-snapshot-defined-styles:282: set end of qualityState to {\"quality-state-end\"}"
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
