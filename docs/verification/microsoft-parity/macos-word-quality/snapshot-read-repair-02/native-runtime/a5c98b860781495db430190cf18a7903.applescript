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
set boundDoc to document "document-ff5405fea4934a65abece05841f98af8.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-htaez118/document-ff5405fea4934a65abece05841f98af8.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-htaez118/document-ff5405fea4934a65abece05841f98af8.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-htaez118/document-ff5405fea4934a65abece05841f98af8.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-htaez118/document-ff5405fea4934a65abece05841f98af8.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "even-pages-header-story-hash:0: set qr to get story range boundDoc story type even pages header story"
log probePhase
set qr to get story range boundDoc story type even pages header story
set probePhase to "even-pages-header-story-hash:1: set recoveryTask to current application's NSTask's alloc()'s init()"
log probePhase
set recoveryTask to current application's NSTask's alloc()'s init()
set probePhase to "even-pages-header-story-hash:2: recoveryTask's setLaunchPath:\"/usr/bin/shasum\""
log probePhase
recoveryTask's setLaunchPath:"/usr/bin/shasum"
set probePhase to "even-pages-header-story-hash:3: recoveryTask's setArguments:{\"-a\", \"256\"}"
log probePhase
recoveryTask's setArguments:{"-a", "256"}
set probePhase to "even-pages-header-story-hash:4: set recoveryInput to current application's NSPipe's pipe()"
log probePhase
set recoveryInput to current application's NSPipe's pipe()
set probePhase to "even-pages-header-story-hash:5: set recoveryOutput to current application's NSPipe's pipe()"
log probePhase
set recoveryOutput to current application's NSPipe's pipe()
set probePhase to "even-pages-header-story-hash:6: recoveryTask's setStandardInput:recoveryInput"
log probePhase
recoveryTask's setStandardInput:recoveryInput
set probePhase to "even-pages-header-story-hash:7: recoveryTask's setStandardOutput:recoveryOutput"
log probePhase
recoveryTask's setStandardOutput:recoveryOutput
set probePhase to "even-pages-header-story-hash:8: recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())"
log probePhase
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set probePhase to "even-pages-header-story-hash:9: set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4"
log probePhase
set recoveryData to (current application's NSString's stringWithString:(content of qr as text))'s dataUsingEncoding:4
set probePhase to "even-pages-header-story-hash:10: recoveryTask's |launch|()"
log probePhase
recoveryTask's |launch|()
set probePhase to "even-pages-header-story-hash:11: (recoveryInput's fileHandleForWriting())'s writeData:recoveryData"
log probePhase
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
set probePhase to "even-pages-header-story-hash:12: (recoveryInput's fileHandleForWriting())'s closeFile()"
log probePhase
(recoveryInput's fileHandleForWriting())'s closeFile()
set probePhase to "even-pages-header-story-hash:13: set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()"
log probePhase
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
set probePhase to "even-pages-header-story-hash:14: recoveryTask's waitUntilExit()"
log probePhase
recoveryTask's waitUntilExit()
set probePhase to "even-pages-header-story-hash:15: if (recoveryTask's terminationStatus() as integer) is not 0 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "even-pages-header-story-hash:16: set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text"
log probePhase
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
set probePhase to "even-pages-header-story-hash:17: if (length of recoveryDigestText) is not 68 then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "even-pages-header-story-hash:18: if text 65 thru 68 of recoveryDigestText is not \"  -\" & linefeed then error \"WPSC_CHECKPOINT_HASH_FAILED\""
log probePhase
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set probePhase to "even-pages-header-story-hash:19: set probePartHash to text 1 thru 64 of recoveryDigestText"
log probePhase
set probePartHash to text 1 thru 64 of recoveryDigestText
set probePhase to "even-pages-header-story-hash:20: set probeRows to {{\"story-hash\",probePartHash,(count fields of qr)}}"
log probePhase
set probeRows to {{"story-hash",probePartHash,(count fields of qr)}}
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
