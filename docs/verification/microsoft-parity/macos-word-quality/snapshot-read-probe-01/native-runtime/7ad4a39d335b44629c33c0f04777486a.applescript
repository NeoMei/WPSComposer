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
set probePhase to "0: set qualityLayout to {}"
set qualityLayout to {}
repeat with qi from 1 to count sections of boundDoc
set probePhase to "2: set qsection to section qi of boundDoc"
set qsection to section qi of boundDoc
set probePhase to "3: set qsetup to page setup of qsection"
set qsetup to page setup of qsection
set probePhase to "4: set end of qualityLayout to {\"section\",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,head"
set end of qualityLayout to {"section",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,header distance of qsetup,footer distance of qsetup,gutter of qsetup,count text columns of qsetup}
repeat with qindex in {header footer primary,header footer first page,header footer even pages}
set probePhase to "6: set qheader to get header qsection index qindex"
set qheader to get header qsection index qindex
set probePhase to "7: set qfooter to get footer qsection index qindex"
set qfooter to get footer qsection index qindex
repeat with qpart in {qheader,qfooter}
set probePhase to "9: if (count shapes of qpart) is not 0 then error \"WPSC_QUALITY_DRAWING_UNVERIFIED\""
if (count shapes of qpart) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"
set probePhase to "10: set qr to text object of qpart"
set qr to text object of qpart
set probePhase to "11: set end of qualityLayout to {\"page-part\",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as text}"
set end of qualityLayout to {"page-part",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as text}
repeat with qfield in fields of qr
set probePhase to "13: set end of qualityLayout to {\"page-field\",field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield,end of content of field co"
set end of qualityLayout to {"page-field",field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield,end of content of field code of qfield,start of content of result range of qfield,end of content of result range of qfield,locked of qfield}
end repeat
end repeat
end repeat
end repeat
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(my jsonRows(qualityLayout)))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probeHash to text 1 thru 64 of recoveryDigestText
set nativeRows to {{"read-ok",count qualityLayout,probeHash}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
