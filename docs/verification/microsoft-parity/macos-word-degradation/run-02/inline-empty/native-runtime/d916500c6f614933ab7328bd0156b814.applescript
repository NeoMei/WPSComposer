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
set boundDoc to document "document-f45bf053a2164841beb3513f5b4787d4.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-7_r6v0og/document-f45bf053a2164841beb3513f5b4787d4.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
set noticeBeforeEnd to end of content of text object of boundDoc
set noticeBeforeTables to count tables of boundDoc
set noticeStart to insertionPoint
set noticeRange to create range boundDoc start noticeStart end noticeStart
set content of noticeRange to "[NOTICE_INLINE_EMPTY: 中文😀]"
set noticeEnd to noticeStart + 27
set noticeRange to create range boundDoc start noticeStart end noticeEnd
set italic of font object of noticeRange to true
set color of font object of noticeRange to {40092, 0, 1542}
set background pattern color of shading of noticeRange to {64764, 59624, 59110}
set nativeRows to {{"degradation-range","inline",start of content of noticeRange,end of content of noticeRange,content of noticeRange as text,noticeBeforeEnd,end of content of text object of boundDoc,noticeBeforeTables,count tables of boundDoc}}
set end of nativeRows to {"degradation-style",italic of font object of noticeRange,(color of font object of noticeRange is {40092, 0, 1542}),(background pattern color of shading of noticeRange is {64764, 59624, 59110})}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
