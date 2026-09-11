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
set boundDoc to document "document-22111abb5db84f10a3267bdfcf23e892.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-1qyy66n5/document-22111abb5db84f10a3267bdfcf23e892.docx" then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-1qyy66n5/document-22111abb5db84f10a3267bdfcf23e892.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-1qyy66n5/document-22111abb5db84f10a3267bdfcf23e892.docx") then error "QUALITY_WINDOW_CHANGED"
activate object boundWindow
set content of text object of boundDoc to "PREFIX 中文😀" & return & "REPLACE" & return & "SUFFIX preserved 中文😀" & return & "OLD TABLE SLOT" & return & ""
set orientation of page setup of boundDoc to orient portrait
set page width of page setup of boundDoc to 600
set page height of page setup of boundDoc to 780
set left margin of page setup of boundDoc to 60
set right margin of page setup of boundDoc to 60
set top margin of page setup of boundDoc to 60
set bottom margin of page setup of boundDoc to 60
set fixtureRange to text object of paragraph 2 of boundDoc
set fixtureRange to create range boundDoc start (start of content of fixtureRange) end ((end of content of fixtureRange) - 1)
make new bookmark at boundDoc with properties {name:"semantic_replace",text object:fixtureRange}
make new bookmark at boundDoc with properties {name:"semantic_suffix",text object:text object of paragraph 3 of boundDoc}
set fixtureRange to text object of paragraph 4 of boundDoc
set fixtureRange to create range boundDoc start (start of content of fixtureRange) end ((end of content of fixtureRange) - 1)
set fixtureOldTable to make new table at boundDoc with properties {text object:fixtureRange,number of rows:1,number of columns:1}
set content of text object of (get cell from table fixtureOldTable row 1 column 1) to "EXISTING TABLE 原样"
make new bookmark at boundDoc with properties {name:"semantic_old_table",text object:text object of fixtureOldTable}
set fixtureRange to text object of bookmark "semantic_replace" of boundDoc
set selection start of selection of boundWindow to start of content of fixtureRange
set selection end of selection of boundWindow to end of content of fixtureRange
set nativeRows to {{"seed",start of content of fixtureRange,end of content of fixtureRange,count tables of boundDoc,content of text object of bookmark "semantic_old_table" of boundDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
