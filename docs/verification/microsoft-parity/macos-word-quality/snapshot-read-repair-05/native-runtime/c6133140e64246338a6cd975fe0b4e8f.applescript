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
set boundDoc to document "document-25275478f9cd4ab8bba2e6f00bf5022e.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-r14ji_fl/document-25275478f9cd4ab8bba2e6f00bf5022e.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-r14ji_fl/document-25275478f9cd4ab8bba2e6f00bf5022e.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-r14ji_fl/document-25275478f9cd4ab8bba2e6f00bf5022e.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-r14ji_fl/document-25275478f9cd4ab8bba2e6f00bf5022e.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "full-snapshot-candidate:0: set qe to end of content of text object of boundDoc"
log probePhase
set qe to end of content of text object of boundDoc
set probePhase to "full-snapshot-candidate:1: set qp to create range boundDoc start 0 end qualityPoint"
log probePhase
set qp to create range boundDoc start 0 end qualityPoint
set probePhase to "full-snapshot-candidate:2: set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe"
log probePhase
set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe
set probePhase to "full-snapshot-candidate:3: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:4: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:5: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:6: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:7: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:8: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:9: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:10: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:11: set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qp as text))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:12: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:13: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:14: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:15: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:16: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:17: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:18: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:19: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:20: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:21: set qualityPrefixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityPrefixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:22: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:23: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:24: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:25: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:26: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:27: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:28: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:29: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:30: set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qs as text))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:31: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:32: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:33: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:34: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:35: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:36: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:37: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:38: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:39: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:40: set qualitySuffixHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualitySuffixHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:41: set qualityState to {{\"quality-state\",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}"
log probePhase
set qualityState to {{"quality-state",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}
set probePhase to "full-snapshot-candidate:42: set qualityLayout to {}"
log probePhase
set qualityLayout to {}
set probePhase to "full-snapshot-candidate:43: repeat with qi from 1 to count sections of boundDoc"
log probePhase
repeat with qi from 1 to count sections of boundDoc
set probePhase to "full-snapshot-candidate:44: set qsection to section qi of boundDoc"
log probePhase
set qsection to section qi of boundDoc
set probePhase to "full-snapshot-candidate:45: set qsetup to page setup of qsection"
log probePhase
set qsetup to page setup of qsection
set probePhase to "full-snapshot-candidate:46: set end of qualityLayout to {\"section\",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of "
log probePhase
set end of qualityLayout to {"section",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,header distance of qsetup,footer distance of qsetup,gutter of qsetup,count text columns of qsetup}
set probePhase to "full-snapshot-candidate:47: repeat with qindex in {header footer primary,header footer first page,header footer even pages}"
log probePhase
repeat with qindex in {header footer primary,header footer first page,header footer even pages}
set probePhase to "full-snapshot-candidate:48: set qheader to get header qsection index qindex"
log probePhase
set qheader to get header qsection index qindex
set probePhase to "full-snapshot-candidate:49: set qfooter to get footer qsection index qindex"
log probePhase
set qfooter to get footer qsection index qindex
set probePhase to "full-snapshot-candidate:50: repeat with qpart in {qheader,qfooter}"
log probePhase
repeat with qpart in {qheader,qfooter}
set probePhase to "full-snapshot-candidate:51: if (count shapes of qpart) is not 0 then error \"WPSC_QUALITY_DRAWING_UNVERIFIED\""
log probePhase
if (count shapes of qpart) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"
set probePhase to "full-snapshot-candidate:52: set end of qualityLayout to {\"page-part-metadata\",qi as integer,header footer index of qpart as text,is header of qpart as boolean,link to p"
log probePhase
set end of qualityLayout to {"page-part-metadata",qi as integer,header footer index of qpart as text,is header of qpart as boolean,link to previous of qpart as boolean}
end repeat
end repeat
end repeat
set probePhase to "full-snapshot-candidate:56: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "full-snapshot-candidate:57: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:58: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:59: set probeStoryRange to get story range boundDoc story type primary header story"
log probePhase
set probeStoryRange to get story range boundDoc story type primary header story
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:61: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:63: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:64: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:66: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:68: set end of qualityLayout to {\"page-story-available\",\"primary header\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","primary header",probeStoryAvailable as boolean}
set probePhase to "full-snapshot-candidate:69: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "full-snapshot-candidate:70: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "full-snapshot-candidate:71: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "full-snapshot-candidate:72: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "full-snapshot-candidate:73: set end of qualityLayout to {\"page-story\",\"primary header\",probeStoryOrdinal as integer,content of probeStoryRange as text}"
log probePhase
set end of qualityLayout to {"page-story","primary header",probeStoryOrdinal as integer,content of probeStoryRange as text}
set probePhase to "full-snapshot-candidate:74: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "full-snapshot-candidate:75: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "full-snapshot-candidate:76: set end of qualityLayout to {\"page-field\",\"primary header\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,c"
log probePhase
set end of qualityLayout to {"page-field","primary header",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "full-snapshot-candidate:78: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:79: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:80: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:82: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:84: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:85: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:87: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "full-snapshot-candidate:90: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "full-snapshot-candidate:91: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:92: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:93: set probeStoryRange to get story range boundDoc story type primary footer story"
log probePhase
set probeStoryRange to get story range boundDoc story type primary footer story
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:95: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:97: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:98: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:100: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:102: set end of qualityLayout to {\"page-story-available\",\"primary footer\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","primary footer",probeStoryAvailable as boolean}
set probePhase to "full-snapshot-candidate:103: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "full-snapshot-candidate:104: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "full-snapshot-candidate:105: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "full-snapshot-candidate:106: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "full-snapshot-candidate:107: set end of qualityLayout to {\"page-story\",\"primary footer\",probeStoryOrdinal as integer,content of probeStoryRange as text}"
log probePhase
set end of qualityLayout to {"page-story","primary footer",probeStoryOrdinal as integer,content of probeStoryRange as text}
set probePhase to "full-snapshot-candidate:108: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "full-snapshot-candidate:109: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "full-snapshot-candidate:110: set end of qualityLayout to {\"page-field\",\"primary footer\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,c"
log probePhase
set end of qualityLayout to {"page-field","primary footer",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "full-snapshot-candidate:112: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:113: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:114: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:116: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:118: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:119: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:121: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "full-snapshot-candidate:124: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "full-snapshot-candidate:125: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:126: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:127: set probeStoryRange to get story range boundDoc story type first page header story"
log probePhase
set probeStoryRange to get story range boundDoc story type first page header story
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:129: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:131: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:132: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:134: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:136: set end of qualityLayout to {\"page-story-available\",\"first page header\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","first page header",probeStoryAvailable as boolean}
set probePhase to "full-snapshot-candidate:137: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "full-snapshot-candidate:138: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "full-snapshot-candidate:139: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "full-snapshot-candidate:140: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "full-snapshot-candidate:141: set end of qualityLayout to {\"page-story\",\"first page header\",probeStoryOrdinal as integer,content of probeStoryRange as text}"
log probePhase
set end of qualityLayout to {"page-story","first page header",probeStoryOrdinal as integer,content of probeStoryRange as text}
set probePhase to "full-snapshot-candidate:142: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "full-snapshot-candidate:143: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "full-snapshot-candidate:144: set end of qualityLayout to {\"page-field\",\"first page header\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","first page header",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "full-snapshot-candidate:146: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:147: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:148: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:150: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:152: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:153: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:155: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "full-snapshot-candidate:158: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "full-snapshot-candidate:159: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:160: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:161: set probeStoryRange to get story range boundDoc story type first page footer story"
log probePhase
set probeStoryRange to get story range boundDoc story type first page footer story
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:163: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:165: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:166: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:168: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:170: set end of qualityLayout to {\"page-story-available\",\"first page footer\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","first page footer",probeStoryAvailable as boolean}
set probePhase to "full-snapshot-candidate:171: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "full-snapshot-candidate:172: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "full-snapshot-candidate:173: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "full-snapshot-candidate:174: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "full-snapshot-candidate:175: set end of qualityLayout to {\"page-story\",\"first page footer\",probeStoryOrdinal as integer,content of probeStoryRange as text}"
log probePhase
set end of qualityLayout to {"page-story","first page footer",probeStoryOrdinal as integer,content of probeStoryRange as text}
set probePhase to "full-snapshot-candidate:176: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "full-snapshot-candidate:177: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "full-snapshot-candidate:178: set end of qualityLayout to {\"page-field\",\"first page footer\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","first page footer",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "full-snapshot-candidate:180: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:181: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:182: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:184: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:186: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:187: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:189: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "full-snapshot-candidate:192: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "full-snapshot-candidate:193: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:194: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:195: set probeStoryRange to get story range boundDoc story type even pages header story"
log probePhase
set probeStoryRange to get story range boundDoc story type even pages header story
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:197: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:199: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:200: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:202: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:204: set end of qualityLayout to {\"page-story-available\",\"even pages header\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","even pages header",probeStoryAvailable as boolean}
set probePhase to "full-snapshot-candidate:205: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "full-snapshot-candidate:206: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "full-snapshot-candidate:207: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "full-snapshot-candidate:208: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "full-snapshot-candidate:209: set end of qualityLayout to {\"page-story\",\"even pages header\",probeStoryOrdinal as integer,content of probeStoryRange as text}"
log probePhase
set end of qualityLayout to {"page-story","even pages header",probeStoryOrdinal as integer,content of probeStoryRange as text}
set probePhase to "full-snapshot-candidate:210: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "full-snapshot-candidate:211: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "full-snapshot-candidate:212: set end of qualityLayout to {\"page-field\",\"even pages header\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","even pages header",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "full-snapshot-candidate:214: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:215: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:216: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:218: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:220: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:221: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:223: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "full-snapshot-candidate:226: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "full-snapshot-candidate:227: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:228: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:229: set probeStoryRange to get story range boundDoc story type even pages footer story"
log probePhase
set probeStoryRange to get story range boundDoc story type even pages footer story
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:231: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:233: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:234: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:236: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:238: set end of qualityLayout to {\"page-story-available\",\"even pages footer\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","even pages footer",probeStoryAvailable as boolean}
set probePhase to "full-snapshot-candidate:239: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "full-snapshot-candidate:240: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "full-snapshot-candidate:241: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "full-snapshot-candidate:242: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "full-snapshot-candidate:243: set end of qualityLayout to {\"page-story\",\"even pages footer\",probeStoryOrdinal as integer,content of probeStoryRange as text}"
log probePhase
set end of qualityLayout to {"page-story","even pages footer",probeStoryOrdinal as integer,content of probeStoryRange as text}
set probePhase to "full-snapshot-candidate:244: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "full-snapshot-candidate:245: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "full-snapshot-candidate:246: set end of qualityLayout to {\"page-field\",\"even pages footer\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","even pages footer",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "full-snapshot-candidate:248: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "full-snapshot-candidate:249: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:250: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:252: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "full-snapshot-candidate:254: try"
log probePhase
try
set probePhase to "full-snapshot-candidate:255: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "full-snapshot-candidate:257: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "full-snapshot-candidate:260: repeat with qi from 1 to count list templates of boundDoc"
log probePhase
repeat with qi from 1 to count list templates of boundDoc
set probePhase to "full-snapshot-candidate:261: set qtemplate to list template qi of boundDoc"
log probePhase
set qtemplate to list template qi of boundDoc
set probePhase to "full-snapshot-candidate:262: set end of qualityLayout to {\"list-template\",qi as integer,name of qtemplate as text,outline numbered of qtemplate}"
log probePhase
set end of qualityLayout to {"list-template",qi as integer,name of qtemplate as text,outline numbered of qtemplate}
set probePhase to "full-snapshot-candidate:263: repeat with qlevel in list levels of qtemplate"
log probePhase
repeat with qlevel in list levels of qtemplate
set probePhase to "full-snapshot-candidate:264: set end of qualityLayout to {\"list-level\",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style "
log probePhase
set end of qualityLayout to {"list-level",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style of qlevel as text,start at of qlevel,reset on higher of qlevel,number position of qlevel,text position of qlevel,tab position of qlevel,trailing character of qlevel as text,list level alignment of qlevel as text}
end repeat
end repeat
set probePhase to "full-snapshot-candidate:267: set qualityStyleNames to get name local of every Word style of boundDoc"
log probePhase
set qualityStyleNames to get name local of every Word style of boundDoc
set probePhase to "full-snapshot-candidate:268: set qualityStyleInUseFlags to get in use of every Word style of boundDoc"
log probePhase
set qualityStyleInUseFlags to get in use of every Word style of boundDoc
set probePhase to "full-snapshot-candidate:269: set qualityStyleBuiltinFlags to get built in of every Word style of boundDoc"
log probePhase
set qualityStyleBuiltinFlags to get built in of every Word style of boundDoc
set probePhase to "full-snapshot-candidate:270: if (count qualityStyleNames) is not (count qualityStyleInUseFlags) or (count qualityStyleNames) is not (count qualityStyleBuiltinFlags) then"
log probePhase
if (count qualityStyleNames) is not (count qualityStyleInUseFlags) or (count qualityStyleNames) is not (count qualityStyleBuiltinFlags) then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set probePhase to "full-snapshot-candidate:271: repeat with qualityStyleOrdinal from 1 to count qualityStyleNames"
log probePhase
repeat with qualityStyleOrdinal from 1 to count qualityStyleNames
set probePhase to "full-snapshot-candidate:272: set qualityStyleCurrentName to item qualityStyleOrdinal of qualityStyleNames as text"
log probePhase
set qualityStyleCurrentName to item qualityStyleOrdinal of qualityStyleNames as text
set probePhase to "full-snapshot-candidate:273: set qualityStyleIsInUse to item qualityStyleOrdinal of qualityStyleInUseFlags"
log probePhase
set qualityStyleIsInUse to item qualityStyleOrdinal of qualityStyleInUseFlags
set probePhase to "full-snapshot-candidate:274: set qualityStyleIsBuiltin to item qualityStyleOrdinal of qualityStyleBuiltinFlags"
log probePhase
set qualityStyleIsBuiltin to item qualityStyleOrdinal of qualityStyleBuiltinFlags
set probePhase to "full-snapshot-candidate:275: set end of qualityLayout to {\"style-state\",qualityStyleCurrentName,qualityStyleIsInUse,qualityStyleIsBuiltin}"
log probePhase
set end of qualityLayout to {"style-state",qualityStyleCurrentName,qualityStyleIsInUse,qualityStyleIsBuiltin}
set probePhase to "full-snapshot-candidate:276: if qualityStyleIsInUse or not qualityStyleIsBuiltin then"
log probePhase
if qualityStyleIsInUse or not qualityStyleIsBuiltin then
set probePhase to "full-snapshot-candidate:277: set qstyle to Word style qualityStyleCurrentName of boundDoc"
log probePhase
set qstyle to Word style qualityStyleCurrentName of boundDoc
set probePhase to "full-snapshot-candidate:278: if (name local of qstyle as text) is not qualityStyleCurrentName then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\""
log probePhase
if (name local of qstyle as text) is not qualityStyleCurrentName then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set probePhase to "full-snapshot-candidate:279: set end of qualityLayout to {\"style\",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}"
log probePhase
set end of qualityLayout to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}
end if
end repeat
set probePhase to "full-snapshot-candidate:282: if (get name local of every Word style of boundDoc) is not qualityStyleNames then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\""
log probePhase
if (get name local of every Word style of boundDoc) is not qualityStyleNames then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set probePhase to "full-snapshot-candidate:283: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:284: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:285: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:286: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:287: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:288: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:289: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:290: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:291: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:292: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:293: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:294: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:295: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:296: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:297: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:298: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:299: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:300: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:301: set qualityLayoutHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityLayoutHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:302: set end of qualityState to {\"layout\",qualityLayoutHash}"
log probePhase
set end of qualityState to {"layout",qualityLayoutHash}
set probePhase to "full-snapshot-candidate:303: repeat with qi from 1 to count paragraphs of boundDoc"
log probePhase
repeat with qi from 1 to count paragraphs of boundDoc
set probePhase to "full-snapshot-candidate:304: set qr to text object of paragraph qi of boundDoc"
log probePhase
set qr to text object of paragraph qi of boundDoc
set probePhase to "full-snapshot-candidate:305: set qa to start of content of qr"
log probePhase
set qa to start of content of qr
set probePhase to "full-snapshot-candidate:306: set qz to end of content of qr"
log probePhase
set qz to end of content of qr
set probePhase to "full-snapshot-candidate:307: if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then"
log probePhase
if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then
set probePhase to "full-snapshot-candidate:308: if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then"
log probePhase
if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then
set probePhase to "full-snapshot-candidate:309: set qa to qualityPoint + qualityDelta"
log probePhase
set qa to qualityPoint + qualityDelta
set probePhase to "full-snapshot-candidate:310: set qr to create range boundDoc start qa end qz"
log probePhase
set qr to create range boundDoc start qa end qz
end if
set probePhase to "full-snapshot-candidate:312: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:313: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:314: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:315: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:316: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:317: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:318: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:319: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:320: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:321: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:322: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:323: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:324: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:325: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:326: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:327: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:328: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:329: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:330: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:331: set qpf to paragraph format of qr"
log probePhase
set qpf to paragraph format of qr
set probePhase to "full-snapshot-candidate:332: set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format righ"
log probePhase
set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,space before of qpf,space after of qpf,line spacing of qpf,line spacing rule of qpf as text,alignment of qpf as text,keep with next of qpf,keep together of qpf,widow control of qpf,outline level of qpf as text,character unit first line indent of qpf,list type of list format of qr as text,list level number of list format of qr,list value of list format of qr,list string of list format of qr as text}}
set probePhase to "full-snapshot-candidate:333: repeat with qc from qa to qz - 1"
log probePhase
repeat with qc from qa to qz - 1
set probePhase to "full-snapshot-candidate:334: set qcr to create range boundDoc start qc end (qc + 1)"
log probePhase
set qcr to create range boundDoc start qc end (qc + 1)
set probePhase to "full-snapshot-candidate:335: set qcf to font object of qcr"
log probePhase
set qcf to font object of qcr
set probePhase to "full-snapshot-candidate:336: set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,backgroun"
log probePhase
set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,color of qcf,background pattern color of shading of qcr}
end repeat
set probePhase to "full-snapshot-candidate:338: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:339: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:340: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:341: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:342: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:343: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:344: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:345: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:346: set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityFormats)))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:347: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:348: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:349: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:350: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:351: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:352: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:353: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:354: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:355: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:356: set qualityFormatHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityFormatHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:357: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
log probePhase
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "full-snapshot-candidate:358: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
log probePhase
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "full-snapshot-candidate:359: set end of qualityState to {\"paragraph\",qa,qz,qualityTextHash,qualityFormatHash}"
log probePhase
set end of qualityState to {"paragraph",qa,qz,qualityTextHash,qualityFormatHash}
end if
end repeat
set probePhase to "full-snapshot-candidate:362: repeat with qi from 1 to count fields of boundDoc"
log probePhase
repeat with qi from 1 to count fields of boundDoc
set probePhase to "full-snapshot-candidate:363: set qf to field qi of boundDoc"
log probePhase
set qf to field qi of boundDoc
set probePhase to "full-snapshot-candidate:364: set qa to start of content of field code of qf"
log probePhase
set qa to start of content of field code of qf
set probePhase to "full-snapshot-candidate:365: set qz to end of content of field code of qf"
log probePhase
set qz to end of content of field code of qf
set probePhase to "full-snapshot-candidate:366: set qra to start of content of result range of qf"
log probePhase
set qra to start of content of result range of qf
set probePhase to "full-snapshot-candidate:367: set qrz to end of content of result range of qf"
log probePhase
set qrz to end of content of result range of qf
set probePhase to "full-snapshot-candidate:368: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:369: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:370: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:371: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:372: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:373: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:374: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:375: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:376: set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of field code of qf as text))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:377: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:378: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:379: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:380: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:381: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:382: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:383: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:384: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:385: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:386: set qualityCodeHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityCodeHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:387: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:388: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:389: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:390: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:391: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:392: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:393: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:394: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:395: set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of result range of qf as text))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:396: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:397: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:398: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:399: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:400: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:401: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:402: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:403: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:404: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:405: set qualityResultHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityResultHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:406: if qa >= qualityPoint + qualityDelta then"
log probePhase
if qa >= qualityPoint + qualityDelta then
set probePhase to "full-snapshot-candidate:407: set qa to qa - qualityDelta"
log probePhase
set qa to qa - qualityDelta
set probePhase to "full-snapshot-candidate:408: set qz to qz - qualityDelta"
log probePhase
set qz to qz - qualityDelta
set probePhase to "full-snapshot-candidate:409: set qra to qra - qualityDelta"
log probePhase
set qra to qra - qualityDelta
set probePhase to "full-snapshot-candidate:410: set qrz to qrz - qualityDelta"
log probePhase
set qrz to qrz - qualityDelta
end if
set probePhase to "full-snapshot-candidate:412: set end of qualityState to {\"field\",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}"
log probePhase
set end of qualityState to {"field",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}
end repeat
set probePhase to "full-snapshot-candidate:414: set qualityOldOrdinal to 0"
log probePhase
set qualityOldOrdinal to 0
set probePhase to "full-snapshot-candidate:415: repeat with qi from 1 to count tables of boundDoc"
log probePhase
repeat with qi from 1 to count tables of boundDoc
set probePhase to "full-snapshot-candidate:416: set qt to table qi of boundDoc"
log probePhase
set qt to table qi of boundDoc
set probePhase to "full-snapshot-candidate:417: set qa to start of content of text object of qt"
log probePhase
set qa to start of content of text object of qt
set probePhase to "full-snapshot-candidate:418: set qz to end of content of text object of qt"
log probePhase
set qz to end of content of text object of qt
set probePhase to "full-snapshot-candidate:419: if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then"
log probePhase
if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then
set probePhase to "full-snapshot-candidate:420: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:421: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:422: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:423: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:424: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:425: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:426: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:427: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:428: set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of text object of qt as text))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:429: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:430: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:431: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:432: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:433: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:434: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:435: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:436: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:437: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:438: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:439: if qa >= qualityPoint + qualityDelta then"
log probePhase
if qa >= qualityPoint + qualityDelta then
set probePhase to "full-snapshot-candidate:440: set qa to qa - qualityDelta"
log probePhase
set qa to qa - qualityDelta
set probePhase to "full-snapshot-candidate:441: set qz to qz - qualityDelta"
log probePhase
set qz to qz - qualityDelta
end if
set probePhase to "full-snapshot-candidate:443: set qualityOldOrdinal to qualityOldOrdinal + 1"
log probePhase
set qualityOldOrdinal to qualityOldOrdinal + 1
set probePhase to "full-snapshot-candidate:444: set qualityRowFlags to {}"
log probePhase
set qualityRowFlags to {}
set probePhase to "full-snapshot-candidate:445: repeat with qualityRowOrdinal from 1 to count rows of qt"
log probePhase
repeat with qualityRowOrdinal from 1 to count rows of qt
set probePhase to "full-snapshot-candidate:446: set qrow to row qualityRowOrdinal of qt"
log probePhase
set qrow to row qualityRowOrdinal of qt
set probePhase to "full-snapshot-candidate:447: set end of qualityRowFlags to allow break across pages of qrow"
log probePhase
set end of qualityRowFlags to allow break across pages of qrow
end repeat
set probePhase to "full-snapshot-candidate:449: set end of qualityState to {\"table\",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}"
log probePhase
set end of qualityState to {"table",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}
end if
end repeat
set probePhase to "full-snapshot-candidate:452: repeat with qi from 1 to count bookmarks of boundDoc"
log probePhase
repeat with qi from 1 to count bookmarks of boundDoc
set probePhase to "full-snapshot-candidate:453: set qb to bookmark qi of boundDoc"
log probePhase
set qb to bookmark qi of boundDoc
set probePhase to "full-snapshot-candidate:454: set qa to start of bookmark of qb"
log probePhase
set qa to start of bookmark of qb
set probePhase to "full-snapshot-candidate:455: set qz to end of bookmark of qb"
log probePhase
set qz to end of bookmark of qb
set probePhase to "full-snapshot-candidate:456: set qhashStart to qa"
log probePhase
set qhashStart to qa
set probePhase to "full-snapshot-candidate:457: if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta"
log probePhase
if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta
set probePhase to "full-snapshot-candidate:458: set qr to create range boundDoc start qhashStart end qz"
log probePhase
set qr to create range boundDoc start qhashStart end qz
set probePhase to "full-snapshot-candidate:459: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "full-snapshot-candidate:460: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "full-snapshot-candidate:461: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "full-snapshot-candidate:462: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:463: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "full-snapshot-candidate:464: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "full-snapshot-candidate:465: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "full-snapshot-candidate:466: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "full-snapshot-candidate:467: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "full-snapshot-candidate:468: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "full-snapshot-candidate:469: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "full-snapshot-candidate:470: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "full-snapshot-candidate:471: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "full-snapshot-candidate:472: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "full-snapshot-candidate:473: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:474: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "full-snapshot-candidate:475: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:476: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "full-snapshot-candidate:477: set qualityTextHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set qualityTextHash to text 1 thru 64 of recoveryDigestText
set probePhase to "full-snapshot-candidate:478: if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta"
log probePhase
if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta
set probePhase to "full-snapshot-candidate:479: if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta"
log probePhase
if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta
set probePhase to "full-snapshot-candidate:480: set end of qualityState to {\"bookmark\",name of qb as text,qa,qz,qualityTextHash}"
log probePhase
set end of qualityState to {"bookmark",name of qb as text,qa,qz,qualityTextHash}
end repeat
set probePhase to "full-snapshot-candidate:482: set end of qualityState to {\"quality-state-end\"}"
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
