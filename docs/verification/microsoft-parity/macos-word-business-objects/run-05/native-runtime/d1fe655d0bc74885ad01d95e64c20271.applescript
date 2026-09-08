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
set boundDoc to document "document-77a00d3c690d42fcb60cd1733d903bea.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-e2k8a_fg/document-77a00d3c690d42fcb60cd1733d903bea.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set previousTableCount to count tables of boundDoc
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set insertionRange to create range boundDoc start insertionPoint end insertionPoint
set insertedTable to make new table at boundDoc with properties {text object:insertionRange, number of rows:3, number of columns:3}
set allow page breaks of insertedTable to false
set left padding of insertedTable to 4
set right padding of insertedTable to 4
set top padding of insertedTable to 1.5
set bottom padding of insertedTable to 1.5
set allow auto fit of insertedTable to false
set allow break across pages of row 1 of insertedTable to false
set height rule of row 1 of insertedTable to row height auto
set ownCell to get cell from table insertedTable row 1 column 1
set content of text object of ownCell to "Header 中文😀"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 90.0
set bold of font object of text object of ownCell to true
set background pattern color of shading of text object of ownCell to {17476, 29298, 50372}
set color of font object of text object of ownCell to {65535, 65535, 65535}
set ownCell to get cell from table insertedTable row 1 column 2
set content of text object of ownCell to "H2"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 150.0
set bold of font object of text object of ownCell to true
set background pattern color of shading of text object of ownCell to {17476, 29298, 50372}
set color of font object of text object of ownCell to {65535, 65535, 65535}
set ownCell to get cell from table insertedTable row 1 column 3
set content of text object of ownCell to "H3"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 105.0
set bold of font object of text object of ownCell to true
set background pattern color of shading of text object of ownCell to {17476, 29298, 50372}
set color of font object of text object of ownCell to {65535, 65535, 65535}
set allow break across pages of row 2 of insertedTable to false
set height rule of row 2 of insertedTable to row height auto
set ownCell to get cell from table insertedTable row 2 column 1
set content of text object of ownCell to "Left"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 90.0
set ownCell to get cell from table insertedTable row 2 column 2
set content of text object of ownCell to "Center"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 150.0
set ownCell to get cell from table insertedTable row 2 column 3
set content of text object of ownCell to "Right"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 105.0
set allow break across pages of row 3 of insertedTable to false
set height rule of row 3 of insertedTable to row height auto
set ownCell to get cell from table insertedTable row 3 column 1
set content of text object of ownCell to "Band"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 90.0
set background pattern color of shading of text object of ownCell to {62194, 62194, 62194}
set ownCell to get cell from table insertedTable row 3 column 2
set content of text object of ownCell to "Wrap Wrap Wrap Wrap Wrap Wrap Wrap Wrap Wrap Wrap Wrap Wrap "
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 150.0
set background pattern color of shading of text object of ownCell to {62194, 62194, 62194}
set ownCell to get cell from table insertedTable row 3 column 3
set content of text object of ownCell to "Tail"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set font size of font object of text object of ownCell to 10
set east asian name of font object of text object of ownCell to "仿宋"
set ascii name of font object of text object of ownCell to "Times New Roman"
set width of ownCell to 105.0
set background pattern color of shading of text object of ownCell to {62194, 62194, 62194}
set ownBorder to get border insertedTable which border border top
set line style of ownBorder to line style single
set line width of ownBorder to line width25 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border insertedTable which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width25 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border insertedTable which border border left
set line style of ownBorder to line style single
set line width of ownBorder to line width25 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border insertedTable which border border right
set line style of ownBorder to line style single
set line width of ownBorder to line width25 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border insertedTable which border border horizontal
set line style of ownBorder to line style single
set line width of ownBorder to line width25 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border insertedTable which border border vertical
set line style of ownBorder to line style single
set line width of ownBorder to line width25 point
set color of ownBorder to {53456, 53456, 53456}
set heading format of row 1 of insertedTable to true
set tableIndex to count tables of boundDoc
if tableIndex is not previousTableCount + 1 then error "WPSC_TABLE_COUNT_DELTA_FAILED"
set insertedTable to table tableIndex of boundDoc
if (number of rows of insertedTable) is not 3 or (number of columns of insertedTable) is not 3 then error "WPSC_TABLE_READBACK_FAILED"
set nativeRows to {{"created", "table", tableIndex, previousTableCount, number of rows of insertedTable, number of columns of insertedTable, allow auto fit of insertedTable}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
