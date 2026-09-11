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
set boundDoc to document "document-d8b949621df9400fbf9b18058ed1e060.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-28v0sjvx/document-d8b949621df9400fbf9b18058ed1e060.docx" then error "WPSC_STALE_DOCUMENT"
set recoveryEnd to end of content of text object of boundDoc
set recoveryBound to recoveryEnd - 1
set recoveryTarget to 0
if recoveryTarget < 0 or recoveryTarget > recoveryBound then error "WPSC_CHECKPOINT_BOUND_FAILED"
set recoveryPrefix to create range boundDoc start 0 end recoveryTarget
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(content of recoveryPrefix as text))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryPrefixHash to text 1 thru 64 of recoveryDigestText
set recoveryGuardStart to recoveryTarget - 512
if recoveryGuardStart < 0 then set recoveryGuardStart to 0
set recoveryGuard to create range boundDoc start recoveryGuardStart end recoveryTarget
set recoveryParagraph to text object of paragraph (count paragraphs of boundDoc) of boundDoc
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(content of recoveryParagraph as text))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryParagraphHash to text 1 thru 64 of recoveryDigestText
set nativeRows to {{"checkpoint-state",1,recoveryBound,recoveryEnd,count paragraphs of boundDoc,start of content of recoveryParagraph,end of content of recoveryParagraph,recoveryPrefixHash,recoveryGuardStart,recoveryTarget,content of recoveryGuard as text,recoveryParagraphHash}}
repeat with recoveryOrdinal from 1 to count tables of boundDoc
set recoveryObject to table recoveryOrdinal of boundDoc
set end of nativeRows to {"table",recoveryOrdinal as integer,start of content of text object of recoveryObject,end of content of text object of recoveryObject,count rows of recoveryObject,count columns of recoveryObject}
end repeat
repeat with recoveryOrdinal from 1 to count fields of boundDoc
set recoveryObject to field recoveryOrdinal of boundDoc
set end of nativeRows to {"field","main",recoveryOrdinal as integer,field type of recoveryObject as text,content of field code of recoveryObject as text,start of content of field code of recoveryObject,end of content of field code of recoveryObject,start of content of result range of recoveryObject,end of content of result range of recoveryObject}
end repeat
repeat with recoveryOrdinal from 1 to count bookmarks of boundDoc
set recoveryBookmark to bookmark recoveryOrdinal of boundDoc
set recoveryBookmarkName to name of recoveryBookmark as text
if {} does not contain recoveryBookmarkName then
set recoveryIdentity to text object of recoveryBookmark
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(content of recoveryIdentity as text))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryBookmarkHash to text 1 thru 64 of recoveryDigestText
set end of nativeRows to {"bookmark-hash",recoveryBookmarkName,start of content of recoveryIdentity,end of content of recoveryIdentity,recoveryBookmarkHash}
end if
end repeat
set end of nativeRows to {"checkpoint-end"}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
