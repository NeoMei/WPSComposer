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
tell application "Microsoft Word"
set matches to {}
repeat with di from 1 to count documents
set d to document di
if (posix full name of d as text) is "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-b03g3kic/document-23ed51f595644bca902e8076a0007ec9.docx" then set end of matches to di
end repeat
if count matches is not 1 then error "CLEANUP_IDENTITY"
set d to document (item 1 of matches)
if (name of d as text) is not "document-23ed51f595644bca902e8076a0007ec9.docx" then error "CLEANUP_NAME"
if saved of d is not true then error "CLEANUP_SAVED"
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(content of text object of d as text))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set actualHash to text 1 thru 64 of recoveryDigestText
if actualHash is not "5ca17486b93f5dd63d55bdc09cc15dd65b474d0db370fccc97df2e06a5a3f84e" then error "CLEANUP_CONTENT"
close d saving no
set nativeRows to {{"closed","/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-b03g3kic/document-23ed51f595644bca902e8076a0007ec9.docx"}}
end tell
return my jsonRows(nativeRows)