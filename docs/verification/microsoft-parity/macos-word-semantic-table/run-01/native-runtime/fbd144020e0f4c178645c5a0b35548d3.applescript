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
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-1qyy66n5/document-22111abb5db84f10a3267bdfcf23e892.docx") then error "QUALITY_BOUND_PATH_CHANGED"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-1qyy66n5/document-22111abb5db84f10a3267bdfcf23e892.docx") then error "QUALITY_WINDOW_CHANGED"
set fixtureSelection to selection of boundWindow
if story type of fixtureSelection is not main text story then error "SEMANTIC_WRONG_STORY"
set fixtureRange to text object of fixtureSelection
if start of content of fixtureRange is not 12 or end of content of fixtureRange is not 19 then error "SEMANTIC_RANGE_CHANGED"
set fixtureRange to create range boundDoc start 12 end 19
if (content of fixtureRange as text) is not "REPLACE" then error "SEMANTIC_RANGE_TEXT_CHANGED"
if (count sections of boundDoc) is not 1 then error "SEMANTIC_SECTION_CHANGED"
set fixturePage to page setup of section 1 of boundDoc
if orientation of fixturePage is not orient portrait or page width of fixturePage is not 600 or left margin of fixturePage is not 60 or right margin of fixturePage is not 60 then error "SEMANTIC_PAGE_CHANGED"
if (count tables of boundDoc) is not 1 then error "SEMANTIC_TABLE_BASELINE_CHANGED"
set semanticInsertion to create range boundDoc start 12 end 19
set semanticTable to make new table at boundDoc with properties {text object:semanticInsertion, number of rows:4, number of columns:3}
set semanticCell to get cell from table semanticTable row 1 column 1
set content of text object of semanticCell to ("编号 ID")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph left
set semanticCell to get cell from table semanticTable row 1 column 2
set content of text object of semanticCell to ("Narrative 内容")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph center
set semanticCell to get cell from table semanticTable row 1 column 3
set content of text object of semanticCell to ("Flag")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph right
set semanticCell to get cell from table semanticTable row 2 column 1
set content of text object of semanticCell to ("A1")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph left
set semanticCell to get cell from table semanticTable row 2 column 2
set content of text object of semanticCell to ("正文 中文😀")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph center
set semanticCell to get cell from table semanticTable row 2 column 3
set content of text object of semanticCell to ("V-MERGE")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph right
set semanticCell to get cell from table semanticTable row 3 column 1
set content of text object of semanticCell to ("A2")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph left
set semanticCell to get cell from table semanticTable row 3 column 2
set content of text object of semanticCell to ("None")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph center
set semanticCell to get cell from table semanticTable row 3 column 3
set content of text object of semanticCell to ("")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph right
set semanticCell to get cell from table semanticTable row 4 column 1
set content of text object of semanticCell to ("H-MERGE")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph left
set semanticCell to get cell from table semanticTable row 4 column 2
set content of text object of semanticCell to ("")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph center
set semanticCell to get cell from table semanticTable row 4 column 3
set content of text object of semanticCell to ("True")
set first line indent of paragraph format of text object of semanticCell to 2.5
set paragraph format left indent of paragraph format of text object of semanticCell to 0.0
set paragraph format right indent of paragraph format of text object of semanticCell to 0.0
set alignment of paragraph format of text object of semanticCell to align paragraph right
set allow break across pages of row options of semanticTable to false
set heading format of row 1 of semanticTable to true
try
auto fit behavior semanticTable behavior auto fit fixed
on error semanticMessage number semanticError
if semanticError is -1712 or semanticError is -609 or semanticError is -128 then error semanticMessage number semanticError
end try
set allow auto fit of semanticTable to false
try
set preferred width type of semanticTable to preferred width points
set preferred width of semanticTable to 480.0
on error semanticMessage number semanticError
if semanticError is -1712 or semanticError is -609 or semanticError is -128 then error semanticMessage number semanticError
end try
set semanticColumn to column 1 of semanticTable
try
set table item width semanticColumn column width 135.927094400347 ruler style adjust none
on error semanticMessage number semanticError
if semanticError is -1712 or semanticError is -609 or semanticError is -128 then error semanticMessage number semanticError
set width of semanticColumn to 135.927094400347
end try
set semanticColumn to column 2 of semanticTable
try
set table item width semanticColumn column width 220.8142163128442 ruler style adjust none
on error semanticMessage number semanticError
if semanticError is -1712 or semanticError is -609 or semanticError is -128 then error semanticMessage number semanticError
set width of semanticColumn to 220.8142163128442
end try
set semanticColumn to column 3 of semanticTable
try
set table item width semanticColumn column width 123.25868928680886 ruler style adjust none
on error semanticMessage number semanticError
if semanticError is -1712 or semanticError is -609 or semanticError is -128 then error semanticMessage number semanticError
set width of semanticColumn to 123.25868928680886
end try
set semanticBorder to get border semanticTable which border border top
set line style of semanticBorder to line style single
set line width of semanticBorder to line width150 point
set semanticBorder to get border semanticTable which border border left
set line style of semanticBorder to line style none
set semanticBorder to get border semanticTable which border border bottom
set line style of semanticBorder to line style single
set line width of semanticBorder to line width75 point
set semanticBorder to get border semanticTable which border border right
set line style of semanticBorder to line style single
set line width of semanticBorder to line width25 point
set semanticBorder to get border semanticTable which border border horizontal
set line style of semanticBorder to line style none
set semanticBorder to get border semanticTable which border border vertical
set line style of semanticBorder to line style single
set line width of semanticBorder to line width25 point
set semanticBorder to get border row 1 of semanticTable which border border bottom
set line style of semanticBorder to line style single
set line width of semanticBorder to line width75 point
set nativeRows to {{"style",preferred width of semanticTable,allow auto fit of semanticTable,heading format of row 1 of semanticTable,allow break across pages of row options of semanticTable,{width of column 1 of semanticTable,width of column 2 of semanticTable,width of column 3 of semanticTable}}}
set fixtureCell to get cell from table semanticTable row 1 column 1
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",1,1,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 1 column 2
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",1,2,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 1 column 3
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",1,3,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 2 column 1
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",2,1,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 2 column 2
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",2,2,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 2 column 3
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",2,3,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 3 column 1
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",3,1,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 3 column 2
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",3,2,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 3 column 3
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",3,3,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 4 column 1
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",4,1,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 4 column 2
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",4,2,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureCell to get cell from table semanticTable row 4 column 3
set fixtureRange to text object of fixtureCell
set end of nativeRows to {"cell",4,3,content of fixtureRange as text,first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,paragraph format right indent of paragraph format of fixtureRange,my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}
set fixtureBorder to get border semanticTable which border border top
set fixtureWidth to 0
if line style of fixtureBorder is not line style none then
set fixtureWidth to -1
if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25
if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75
if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5
end if
set end of nativeRows to {"border","top",(line style of fixtureBorder is line style single),fixtureWidth}
set fixtureBorder to get border semanticTable which border border left
set fixtureWidth to 0
if line style of fixtureBorder is not line style none then
set fixtureWidth to -1
if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25
if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75
if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5
end if
set end of nativeRows to {"border","left",(line style of fixtureBorder is line style single),fixtureWidth}
set fixtureBorder to get border semanticTable which border border bottom
set fixtureWidth to 0
if line style of fixtureBorder is not line style none then
set fixtureWidth to -1
if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25
if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75
if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5
end if
set end of nativeRows to {"border","bottom",(line style of fixtureBorder is line style single),fixtureWidth}
set fixtureBorder to get border semanticTable which border border right
set fixtureWidth to 0
if line style of fixtureBorder is not line style none then
set fixtureWidth to -1
if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25
if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75
if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5
end if
set end of nativeRows to {"border","right",(line style of fixtureBorder is line style single),fixtureWidth}
set fixtureBorder to get border semanticTable which border border horizontal
set fixtureWidth to 0
if line style of fixtureBorder is not line style none then
set fixtureWidth to -1
if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25
if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75
if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5
end if
set end of nativeRows to {"border","insideHorizontal",(line style of fixtureBorder is line style single),fixtureWidth}
set fixtureBorder to get border semanticTable which border border vertical
set fixtureWidth to 0
if line style of fixtureBorder is not line style none then
set fixtureWidth to -1
if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25
if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75
if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5
end if
set end of nativeRows to {"border","insideVertical",(line style of fixtureBorder is line style single),fixtureWidth}
set fixtureBorder to get border row 1 of semanticTable which border border bottom
set fixtureWidth to 0
if line style of fixtureBorder is not line style none then
set fixtureWidth to -1
if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25
if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75
if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5
end if
set end of nativeRows to {"border","headerBottom",(line style of fixtureBorder is line style single),fixtureWidth}
set semanticMergeStart to get cell from table semanticTable row 2 column 3
set semanticMergeEnd to get cell from table semanticTable row 3 column 3
merge cell semanticMergeStart with semanticMergeEnd
set semanticMergeStart to get cell from table semanticTable row 4 column 1
set semanticMergeEnd to get cell from table semanticTable row 4 column 2
merge cell semanticMergeStart with semanticMergeEnd
set semanticTableEnd to end of content of text object of semanticTable
set semanticSelection to selection of boundWindow
set selection start of semanticSelection to semanticTableEnd
set selection end of semanticSelection to semanticTableEnd
set semanticTail to create range boundDoc start semanticTableEnd end semanticTableEnd
set content of semanticTail to return
set selection start of semanticSelection to semanticTableEnd + 1
set selection end of semanticSelection to semanticTableEnd + 1
make new bookmark at boundDoc with properties {name:"wpsc_semantic_fixture_table",text object:text object of semanticTable}
set fixtureStyleRows to nativeRows
set fixtureBookmark to text object of bookmark "wpsc_semantic_fixture_table" of boundDoc
set fixtureMatches to 0
repeat with fixtureIndex from 1 to count tables of boundDoc
set fixtureCandidate to table fixtureIndex of boundDoc
set fixtureRange to text object of fixtureCandidate
if start of content of fixtureRange is start of content of fixtureBookmark and end of content of fixtureRange is end of content of fixtureBookmark then
set fixtureMatches to fixtureMatches + 1
set fixtureTable to fixtureCandidate
end if
end repeat
if fixtureMatches is not 1 then error "SEMANTIC_TABLE_IDENTITY"
set fixtureEnd to end of content of text object of fixtureTable
set fixtureTail to create range boundDoc start fixtureEnd end (fixtureEnd + 1)
set fixtureSelectionRange to text object of selection of boundWindow
set nativeRows to {{"result",posix full name of boundDoc as text,start of content of text object of fixtureTable,fixtureEnd,end of content of text object of boundDoc,count tables of boundDoc,start of content of fixtureSelectionRange,end of content of fixtureSelectionRange,content of fixtureTail as text,(story type of fixtureSelectionRange is main text story)}}
set fixturePrefix to create range boundDoc start 0 end 12
set fixtureOld to text object of bookmark "semantic_old_table" of boundDoc
set end of nativeRows to {"outside",content of fixturePrefix as text,content of text object of bookmark "semantic_suffix" of boundDoc as text,content of fixtureOld as text,start of content of fixtureOld}
set end of nativeRows to {"table-text",content of text object of fixtureTable as text}
set nativeRows to fixtureStyleRows & nativeRows
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
