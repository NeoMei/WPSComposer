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
set boundDoc to document "document-023aa3676a2f4caa803721d71fbe42ea.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ftr0g2us/document-023aa3676a2f4caa803721d71fbe42ea.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
set noticeBeforeEnd to end of content of text object of boundDoc
set noticeBeforeTables to count tables of boundDoc
if insertionPoint is not 11 then error "WPSC_DEGRADATION_ANCHOR_CHANGED"
set noticeStart to insertionPoint
set noticeRange to create range boundDoc start noticeStart end noticeStart
set content of noticeRange to "[NOTICE_TABLE_FAILURE_NONEMPTY] 中文😀"
set noticeEnd to noticeStart + 36
set noticeRange to create range boundDoc start noticeStart end noticeEnd
set italic of font object of noticeRange to true
set color of font object of noticeRange to {40092, 0, 1542}
set background pattern color of shading of noticeRange to {64764, 59624, 59110}
set space before of paragraph format of noticeRange to 0
set space after of paragraph format of noticeRange to 3
set keep together of paragraph format of noticeRange to true
set outline level of paragraph format of noticeRange to outline level body text
set nativeRows to {{"degradation-range","block-range",start of content of noticeRange,end of content of noticeRange,content of noticeRange as text,noticeBeforeEnd,end of content of text object of boundDoc,noticeBeforeTables,count tables of boundDoc}}
set end of nativeRows to {"degradation-style",italic of font object of noticeRange,(color of font object of noticeRange is {40092, 0, 1542}),(background pattern color of shading of noticeRange is {64764, 59624, 59110})}
set end of nativeRows to {"degradation-paragraph",space before of paragraph format of noticeRange,space after of paragraph format of noticeRange,keep together of paragraph format of noticeRange,(outline level of paragraph format of noticeRange is outline level body text)}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
