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
set boundDoc to document "document-f6a26a40df8a43cc8bce00276d1790c7.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-0wqhf_fp/document-f6a26a40df8a43cc8bce00276d1790c7.docx" then error "WPSC_STALE_DOCUMENT"
set recoveryEnd to end of content of text object of boundDoc
set recoveryBound to recoveryEnd - 1
set recoveryTarget to 246
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
if not (exists bookmark "WPSC_F_add810dec3c341df9324238e8f5f97" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_F_add810dec3c341df9324238e8f5f97" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_F_add810dec3c341df9324238e8f5f97",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_R_ff7bbb69cd9342d3b67179ae6688cd" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_R_ff7bbb69cd9342d3b67179ae6688cd" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_R_ff7bbb69cd9342d3b67179ae6688cd",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
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
if {("WPSC_F_add810dec3c341df9324238e8f5f97"),("WPSC_R_ff7bbb69cd9342d3b67179ae6688cd")} does not contain recoveryBookmarkName then
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
if not ((current application's NSArray's arrayWithArray:nativeRows)'s isEqualToArray:{{("checkpoint-state"),1,246,247,9,246,247,("914227745b353131b3d059030170b61dd39f662c912ed43d49ef554cca26a2f9"),0,246,("Recovery prefix 中文😀" & return & "Chapter" & return & "Contents" & return & "Chapter" & tab & "1" & return & "" & return & "" & (character id 12) & "" & return & "Existing Recovery" & return & "FALLBACK-ONCE" & return & ""),("9d1e0e2d9459d06523ad13e28a4093c2316baafe7aec5b25f30eba2e113599c4")},{("field"),("main"),1,("field toc"),(" TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT "),39,74,75,149},{("field"),("main"),2,("field hyperlink"),(" HYPERLINK \\l \"_Toc239854507\" "),76,107,108,147},{("field"),("main"),3,("field page ref"),(" PAGEREF _Toc239854507 \\h "),117,144,145,146},{("field"),("main"),4,("field ref"),(" REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT "),163,221,222,230},{("bookmark"),("WPSC_F_add810dec3c341df9324238e8f5f97"),39,74,(" TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT ")},{("bookmark"),("WPSC_R_ff7bbb69cd9342d3b67179ae6688cd"),163,221,(" REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT ")},{("bookmark-hash"),("wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa"),0,8,("48f6a8d5688b0cf59fb8109b7903507ed9d2e1580be2ad7ae169df659e1ddeea")},{("objects"),0,0,0,4,3},{("checkpoint-end")}}) then error "WPSC_CHECKPOINT_PREFLIGHT_CHANGED"
set recoveryDeletedFields to {}
set recoveryDeletedTables to {}
set recoveryLiveTables to {}
set nativeRows to {}
repeat with recoveryOrdinal from 1 to count fields of boundDoc
set recoveryObject to field recoveryOrdinal of boundDoc
set end of nativeRows to {"field","main",recoveryOrdinal as integer,field type of recoveryObject as text,content of field code of recoveryObject as text,start of content of field code of recoveryObject,end of content of field code of recoveryObject,start of content of result range of recoveryObject,end of content of result range of recoveryObject}
end repeat
if not ((current application's NSArray's arrayWithArray:nativeRows)'s isEqualToArray:{{("field"),("main"),1,("field toc"),(" TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT "),39,74,75,149},{("field"),("main"),2,("field hyperlink"),(" HYPERLINK \\l \"_Toc239854507\" "),76,107,108,147},{("field"),("main"),3,("field page ref"),(" PAGEREF _Toc239854507 \\h "),117,144,145,146},{("field"),("main"),4,("field ref"),(" REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT "),163,221,222,230}}) then error "WPSC_CHECKPOINT_UNPLANNED_FIELD_REMOVAL"
set recoveryActualTables to {}
repeat with recoveryOrdinal from 1 to count tables of boundDoc
set recoveryTable to table recoveryOrdinal of boundDoc
set end of recoveryActualTables to {"table",recoveryOrdinal as integer,start of content of text object of recoveryTable,end of content of text object of recoveryTable,count rows of recoveryTable,count columns of recoveryTable}
end repeat
if not ((current application's NSArray's arrayWithArray:recoveryActualTables)'s isEqualToArray:recoveryLiveTables) then error "WPSC_CHECKPOINT_TABLE_CHANGED"
set recoveryBound to (end of content of text object of boundDoc) - 1
if recoveryBound < 246 then error "WPSC_CHECKPOINT_BOUND_FAILED"
set rollbackRange to create range boundDoc start 246 end recoveryBound
set content of rollbackRange to ""
set recoveryEnd to end of content of text object of boundDoc
set recoveryBound to recoveryEnd - 1
set recoveryTarget to 246
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
if not (exists bookmark "WPSC_F_add810dec3c341df9324238e8f5f97" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_F_add810dec3c341df9324238e8f5f97" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_F_add810dec3c341df9324238e8f5f97",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
if not (exists bookmark "WPSC_R_ff7bbb69cd9342d3b67179ae6688cd" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set recoveryIdentity to text object of bookmark "WPSC_R_ff7bbb69cd9342d3b67179ae6688cd" of boundDoc
if not ((current application's NSString's stringWithString:(content of recoveryIdentity as text))'s isEqualToString:(" REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT ")) then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"bookmark","WPSC_R_ff7bbb69cd9342d3b67179ae6688cd",start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}
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
if {("WPSC_F_add810dec3c341df9324238e8f5f97"),("WPSC_R_ff7bbb69cd9342d3b67179ae6688cd")} does not contain recoveryBookmarkName then
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
if not ((current application's NSArray's arrayWithArray:nativeRows)'s isEqualToArray:{{("checkpoint-state"),1,246,247,9,246,247,("914227745b353131b3d059030170b61dd39f662c912ed43d49ef554cca26a2f9"),0,246,("Recovery prefix 中文😀" & return & "Chapter" & return & "Contents" & return & "Chapter" & tab & "1" & return & "" & return & "" & (character id 12) & "" & return & "Existing Recovery" & return & "FALLBACK-ONCE" & return & ""),("9d1e0e2d9459d06523ad13e28a4093c2316baafe7aec5b25f30eba2e113599c4")},{("field"),("main"),1,("field toc"),(" TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT "),39,74,75,149},{("field"),("main"),2,("field hyperlink"),(" HYPERLINK \\l \"_Toc239854507\" "),76,107,108,147},{("field"),("main"),3,("field page ref"),(" PAGEREF _Toc239854507 \\h "),117,144,145,146},{("field"),("main"),4,("field ref"),(" REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT "),163,221,222,230},{("bookmark"),("WPSC_F_add810dec3c341df9324238e8f5f97"),39,74,(" TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT ")},{("bookmark"),("WPSC_R_ff7bbb69cd9342d3b67179ae6688cd"),163,221,(" REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT ")},{("bookmark-hash"),("wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa"),0,8,("48f6a8d5688b0cf59fb8109b7903507ed9d2e1580be2ad7ae169df659e1ddeea")},{("objects"),0,0,0,4,3},{("checkpoint-end")}}) then error "WPSC_CHECKPOINT_POSTCONDITION_FAILED"
set nativeRows to {{"rollback-ack",recoveryDeletedFields,recoveryDeletedTables}} & nativeRows
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
