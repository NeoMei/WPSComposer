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
set boundDoc to document "document-651fc661d3a64ce6a966ffc90e2e5750.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-hhk9noow/document-651fc661d3a64ce6a966ffc90e2e5750.docx" then error "WPSC_STALE_DOCUMENT"
set recoveryEnd to end of content of text object of boundDoc
set recoveryBound to recoveryEnd - 1
set recoveryTarget to 413
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
if not (exists bookmark "WPSC_R_e3250ec6fa444ec2960bfda70a6a0f" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_R_e3250ec6fa444ec2960bfda70a6a0f" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" REF wpsc_eq_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_R_e3250ec6fa444ec2960bfda70a6a0f",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_R_b795c73481134233bb2f4832cabe25" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_R_b795c73481134233bb2f4832cabe25" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" REF wpsc_eq_bbbbbbbbbbbbbbbbbbbbbbbb \\h \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_R_b795c73481134233bb2f4832cabe25",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_R_afaa0c6458f1498db6a84fbccbf861" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_R_afaa0c6458f1498db6a84fbccbf861" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" REF wpsc_eq_cccccccccccccccccccccccc \\h \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_R_afaa0c6458f1498db6a84fbccbf861",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_N_f39aad095aaa4987a3b159b5e66ba0" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_N_f39aad095aaa4987a3b159b5e66ba0" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" SEQ WPSC_EQ \\* ARABIC \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_N_f39aad095aaa4987a3b159b5e66ba0",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_N_20674658b43f4b6c8eb8cc3f817a7d" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_N_20674658b43f4b6c8eb8cc3f817a7d" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" STYLEREF \"标题 1\" \\s \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_N_20674658b43f4b6c8eb8cc3f817a7d",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_N_8d9731ad50e8498f9efcd28149a752" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_N_8d9731ad50e8498f9efcd28149a752" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" SEQ WPSC_FIG \\* ARABIC \\s 1 \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_N_8d9731ad50e8498f9efcd28149a752",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_N_09c925f5e5b346faab7977f2afdbf9" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_N_09c925f5e5b346faab7977f2afdbf9" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" SEQ WPSC_EQ \\* ARABIC \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_N_09c925f5e5b346faab7977f2afdbf9",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
repeat with recoveryOrdinal from 1 to count shapes of boundDoc
set recoveryShape to shape recoveryOrdinal of boundDoc
set recoveryShapeText to ""
if has text of text frame of recoveryShape then set recoveryShapeText to content of text range of text frame of recoveryShape as text
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(recoveryShapeText))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryShapeHash to text 1 thru 64 of recoveryDigestText
set end of nativeRows to {"shape",recoveryOrdinal as integer,name of recoveryShape as text,shape type of recoveryShape as text,anchorID of recoveryShape,editID of recoveryShape,start of content of anchor of recoveryShape,end of content of anchor of recoveryShape,left position of recoveryShape,top of recoveryShape,width of recoveryShape,height of recoveryShape,rotation of recoveryShape,z order position of recoveryShape,visible of recoveryShape,recoveryShapeHash}
end repeat
repeat with recoveryOrdinal from 1 to count inline shapes of boundDoc
set recoveryPicture to inline shape recoveryOrdinal of boundDoc
set recoveryTask to current application's NSTask's alloc()'s init()
recoveryTask's setLaunchPath:"/usr/bin/shasum"
recoveryTask's setArguments:{"-a", "256"}
set recoveryInput to current application's NSPipe's pipe()
set recoveryOutput to current application's NSPipe's pipe()
recoveryTask's setStandardInput:recoveryInput
recoveryTask's setStandardOutput:recoveryOutput
recoveryTask's setStandardError:(current application's NSFileHandle's fileHandleWithNullDevice())
set recoveryData to (current application's NSString's stringWithString:(alternative text of recoveryPicture as text))'s dataUsingEncoding:4
recoveryTask's |launch|()
(recoveryInput's fileHandleForWriting())'s writeData:recoveryData
(recoveryInput's fileHandleForWriting())'s closeFile()
set recoveryDigestData to (recoveryOutput's fileHandleForReading())'s readDataToEndOfFile()
recoveryTask's waitUntilExit()
if (recoveryTask's terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryDigestText to (current application's NSString's alloc()'s initWithData:recoveryDigestData |encoding|:4) as text
if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"
if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"
set recoveryPictureHash to text 1 thru 64 of recoveryDigestText
set end of nativeRows to {"inline-shape",recoveryOrdinal as integer,anchorID of recoveryPicture,editID of recoveryPicture,start of content of text object of recoveryPicture,end of content of text object of recoveryPicture,width of recoveryPicture,height of recoveryPicture,inline shape type of recoveryPicture as text,recoveryPictureHash}
end repeat
repeat with recoveryOrdinal from 1 to count bookmarks of boundDoc
set recoveryBookmark to bookmark recoveryOrdinal of boundDoc
set recoveryBookmarkName to name of recoveryBookmark as text
if {("WPSC_R_e3250ec6fa444ec2960bfda70a6a0f"),("WPSC_R_b795c73481134233bb2f4832cabe25"),("WPSC_R_afaa0c6458f1498db6a84fbccbf861"),("WPSC_N_f39aad095aaa4987a3b159b5e66ba0"),("WPSC_N_20674658b43f4b6c8eb8cc3f817a7d"),("WPSC_N_8d9731ad50e8498f9efcd28149a752"),("WPSC_N_09c925f5e5b346faab7977f2afdbf9")} does not contain recoveryBookmarkName then
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
set end of nativeRows to {"objects",count shapes of boundDoc,count inline shapes of boundDoc,count tables of boundDoc,count fields of boundDoc,count bookmarks of boundDoc}
set end of nativeRows to {"checkpoint-end"}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
