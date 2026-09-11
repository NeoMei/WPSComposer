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
set boundDoc to document "document-7c6ba1877ddd4683913fa03a22f2b578.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ly1lxrfe/document-7c6ba1877ddd4683913fa03a22f2b578.docx" then error "WPSC_STALE_DOCUMENT"
set headingRange to create range boundDoc start 126 end 126
set pendingParagraph to paragraph 1 of headingRange
if (start of content of text object of pendingParagraph) is not 126 then error "WPSC_PENDING_HEADING_STALE"
if (end of content of text object of pendingParagraph) is not 133 then error "WPSC_PENDING_HEADING_STALE"
if (name local of style of text object of pendingParagraph as text) is not (name local of Word style (style heading2) of boundDoc as text) then error "WPSC_PENDING_HEADING_STALE"
set previousSectionCount to count sections of boundDoc
insert break at headingRange break type section break next page
if (count sections of boundDoc) is not previousSectionCount + 1 then error "WPSC_SECTION_COUNT_FAILED"
set ownSection to section (count sections of boundDoc) of boundDoc
set orientation of page setup of ownSection to orient landscape
if orientation of page setup of ownSection is not orient landscape then error "WPSC_SECTION_READBACK_FAILED"
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
