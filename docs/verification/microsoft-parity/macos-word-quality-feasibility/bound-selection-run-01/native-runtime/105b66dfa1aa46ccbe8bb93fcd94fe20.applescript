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
set boundDoc to document "document-feb1b048cd484362b253fbbe673f5569.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySentinel to document "文档75"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档75") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"QUALITY SENTINEL 中文😀 693956e92ffb483189a538bca16a7d4b" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
activate object (active window of qualitySentinel)
set activeSentinelVerified to ((current application's NSString's stringWithString:(name of document of active window as text))'s isEqualToString:"文档75") as boolean
if not activeSentinelVerified then error "QUALITY_SENTINEL_NOT_ACTIVE"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx") then error "QUALITY_WINDOW_CHANGED"
set qualitySelection to selection of boundWindow
set qualityRange to text object of qualitySelection
set qualityStart to start of content of qualityRange
set qualityEnd to end of content of qualityRange
set activeRange to text object of selection of active window
set selectionOwned to ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5parjt3t/document-feb1b048cd484362b253fbbe673f5569.docx") as boolean
if not selectionOwned then error "QUALITY_SELECTION_FOREIGN"
if story type of qualitySelection is not main text story then error "QUALITY_SELECTION_WRONG_STORY"
if qualityStart is not 9 or qualityEnd is not 9 then error "QUALITY_SELECTION_MOVED"
set beforeEnd to end of content of text object of boundDoc
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(content of text object of boundDoc as text))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set beforeHash to text 1 thru 64 of recoveryDigestText
set qualityAnchor to create range boundDoc start qualityEnd end qualityEnd
make new bookmark at boundDoc with properties {name:"wpsc_document_quality_anchor",text object:qualityAnchor}
set qualityBookmark to bookmark "wpsc_document_quality_anchor" of boundDoc
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(content of text object of boundDoc as text))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set afterHash to text 1 thru 64 of recoveryDigestText
set nativeRows to {{"selection",posix full name of boundDoc as text,id of boundWindow,id of active window,qualityEnd,start of content of activeRange,end of content of activeRange,beforeEnd,start of bookmark of qualityBookmark,end of bookmark of qualityBookmark,empty of qualityBookmark,selectionOwned,activeSentinelVerified,beforeHash,afterHash}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
