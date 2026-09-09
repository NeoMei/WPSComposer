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
set boundDoc to document "document-fbbfa0415f3049e1961de23d9748580d.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-3l7nvcew/document-fbbfa0415f3049e1961de23d9748580d.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint is not 11 then error "WPSC_DEGRADATION_ANCHOR_CHANGED"
set noticeBeforeEnd to end of content of text object of boundDoc
set noticeBeforeTables to count tables of boundDoc
try
activate object boundWindow
set noticeTarget to create range boundDoc start insertionPoint end insertionPoint
set noticeTable to make new table at boundDoc with properties {text object:noticeTarget,number of rows:1,number of columns:1}
set probeCell to get cell from table noticeTable row 1 column 1
set content of text object of probeCell to "PARTIAL-NOTICE-REMNANT"
error "WPSC_CONTROLLED_NOTICE_TABLE_FAILURE" number -2700
set noticeCell to get cell from table noticeTable row 1 column 1
set content of text object of noticeCell to "[NOTICE_TABLE_FAILURE_NONEMPTY] 中文😀"
set noticeRange to text object of noticeCell
set italic of font object of noticeRange to true
set color of font object of noticeRange to {40092, 0, 1542}
set background pattern color of shading of noticeRange to {64764, 59624, 59110}
set space before of paragraph format of noticeRange to 0
set space after of paragraph format of noticeRange to 3
set keep together of paragraph format of noticeRange to true
set outline level of paragraph format of noticeRange to outline level body text
set allow break across pages of row 1 of noticeTable to false
set noticeStart to start of content of text object of noticeCell
set noticeEnd to noticeStart + 36
set noticeDisplayRange to create range boundDoc start noticeStart end noticeEnd
set nativeRows to {{"degradation-table",11,start of content of text object of noticeTable,end of content of text object of noticeTable,content of text object of noticeTable as text,noticeBeforeEnd,end of content of text object of boundDoc,noticeBeforeTables,count tables of boundDoc,count rows of noticeTable,count columns of noticeTable,noticeStart,noticeEnd,content of noticeDisplayRange as text}}
set end of nativeRows to {"degradation-style",italic of font object of noticeRange,(color of font object of noticeRange is {40092, 0, 1542}),(background pattern color of shading of noticeRange is {64764, 59624, 59110})}
set end of nativeRows to {"degradation-paragraph",space before of paragraph format of noticeRange,space after of paragraph format of noticeRange,keep together of paragraph format of noticeRange,(outline level of paragraph format of noticeRange is outline level body text)}
set end of nativeRows to {"degradation-row",allow break across pages of row 1 of noticeTable}
on error noticeError number noticeNumber
if noticeNumber is -1712 or noticeNumber is -609 or noticeNumber is -128 then error noticeError number noticeNumber
set nativeRows to {{"degradation-table-failed",11}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
