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
set boundDoc to document "document-66802e6857724bf6bdc17919f32b7778.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-2zk_5x3f/document-66802e6857724bf6bdc17919f32b7778.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set table1Range to create range boundDoc start insertionPoint end insertionPoint
set table1 to make new table at boundDoc with properties {text object:table1Range, number of rows:3, number of columns:3}
set allow auto fit of table1 to false
set allow page breaks of table1 to false
set heading format of row 1 of table1 to true
set left padding of table1 to 4
set right padding of table1 to 4
set top padding of table1 to 2
set bottom padding of table1 to 2
set allow break across pages of row 1 of table1 to false
set ownCell to get cell from table table1 row 1 column 1
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
set width of ownCell to {90, 150, 105}'s item 1
set background pattern color of shading of text object of ownCell to {17476, 29041, 50372}
set color of font object of text object of ownCell to {65535, 65535, 65535}
set bold of font object of text object of ownCell to true
set ownCell to get cell from table table1 row 1 column 2
set content of text object of ownCell to "Header ASCII"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to {90, 150, 105}'s item 2
set background pattern color of shading of text object of ownCell to {17476, 29041, 50372}
set color of font object of text object of ownCell to {65535, 65535, 65535}
set bold of font object of text object of ownCell to true
set ownCell to get cell from table table1 row 1 column 3
set content of text object of ownCell to "Header Three"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to {90, 150, 105}'s item 3
set background pattern color of shading of text object of ownCell to {17476, 29041, 50372}
set color of font object of text object of ownCell to {65535, 65535, 65535}
set bold of font object of text object of ownCell to true
set allow break across pages of row 2 of table1 to false
set ownCell to get cell from table table1 row 2 column 1
set content of text object of ownCell to "左对齐中文😀"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to {90, 150, 105}'s item 1
set ownCell to get cell from table table1 row 2 column 2
set content of text object of ownCell to "Centered"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to {90, 150, 105}'s item 2
set ownCell to get cell from table table1 row 2 column 3
set content of text object of ownCell to "Right aligned"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to {90, 150, 105}'s item 3
set allow break across pages of row 3 of table1 to false
set ownCell to get cell from table table1 row 3 column 1
set content of text object of ownCell to "Banded row"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to {90, 150, 105}'s item 1
set background pattern color of shading of text object of ownCell to {61680, 63222, 64764}
set ownCell to get cell from table table1 row 3 column 2
set content of text object of ownCell to "Wrap text Wrap text Wrap text Wrap text Wrap text Wrap text Wrap text Wrap text "
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to {90, 150, 105}'s item 2
set background pattern color of shading of text object of ownCell to {61680, 63222, 64764}
set ownCell to get cell from table table1 row 3 column 3
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
set width of ownCell to {90, 150, 105}'s item 3
set background pattern color of shading of text object of ownCell to {61680, 63222, 64764}
set ownBorder to get border table1 which border border top
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border table1 which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border table1 which border border left
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border table1 which border border right
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border table1 which border border horizontal
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {53456, 53456, 53456}
set ownBorder to get border table1 which border border vertical
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {53456, 53456, 53456}
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set table2Range to create range boundDoc start insertionPoint end insertionPoint
set table2 to make new table at boundDoc with properties {text object:table2Range, number of rows:3, number of columns:3}
set allow auto fit of table2 to false
set allow break across pages of row 1 of table2 to false
set ownCell to get cell from table table2 row 1 column 1
set content of text object of ownCell to "Horizontal merge"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set ownCell to get cell from table table2 row 1 column 2
set content of text object of ownCell to ""
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set ownCell to get cell from table table2 row 1 column 3
set content of text object of ownCell to "M1 outside"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set allow break across pages of row 2 of table2 to false
set ownCell to get cell from table table2 row 2 column 1
set content of text object of ownCell to "M2 outside"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set ownCell to get cell from table table2 row 2 column 2
set content of text object of ownCell to "M2 center"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set ownCell to get cell from table table2 row 2 column 3
set content of text object of ownCell to "Vertical merge"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set allow break across pages of row 3 of table2 to false
set ownCell to get cell from table table2 row 3 column 1
set content of text object of ownCell to "M3 outside"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph left
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set ownCell to get cell from table table2 row 3 column 2
set content of text object of ownCell to "M3 center"
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph center
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set ownCell to get cell from table table2 row 3 column 3
set content of text object of ownCell to ""
set first line indent of paragraph format of text object of ownCell to 0
set character unit first line indent of paragraph format of text object of ownCell to 0
set paragraph format left indent of paragraph format of text object of ownCell to 0
set paragraph format right indent of paragraph format of text object of ownCell to 0
set space before of paragraph format of text object of ownCell to 0
set space after of paragraph format of text object of ownCell to 0
set line spacing rule of paragraph format of text object of ownCell to line space single
set alignment of paragraph format of text object of ownCell to align paragraph right
set vertical alignment of ownCell to cell align vertical center
set width of ownCell to 110
set mergeStart to get cell from table table2 row 2 column 3
set mergeEnd to get cell from table table2 row 3 column 3
merge cell mergeStart with mergeEnd
set mergeStart to get cell from table table2 row 1 column 1
set mergeEnd to get cell from table table2 row 1 column 2
merge cell mergeStart with mergeEnd
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set inline1Range to create range boundDoc start insertionPoint end insertionPoint
set inline1 to make new inline picture at inline1Range with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-2zk_5x3f/probe-image.png", link to file:false, save with document:true}
set inline1 to inline picture (count inline pictures of boundDoc) of boundDoc
set lock aspect ratio of inline1 to true
set width of inline1 to 72
set alternative text of inline1 to "Inline 中文😀 alt"
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set inline2Range to create range boundDoc start insertionPoint end insertionPoint
set inline2 to make new inline picture at inline2Range with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-2zk_5x3f/probe-image.png", link to file:false, save with document:true}
set inline2 to inline picture (count inline pictures of boundDoc) of boundDoc
set lock aspect ratio of inline2 to false
set width of inline2 to 90
set height of inline2 to 45
set alternative text of inline2 to "Bounded two dimension alt"
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set floatingRange to create range boundDoc start insertionPoint end insertionPoint
set floatingPicture to make new picture at boundDoc with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-2zk_5x3f/probe-image.png", link to file:false, save with document:true, anchor:floatingRange}
set floatingPicture to shape (count shapes of boundDoc) of boundDoc
set lock aspect ratio of floatingPicture to false
set left position of floatingPicture to 300
set top of floatingPicture to 120
set width of floatingPicture to 96
set height of floatingPicture to 54
set wrap type of wrap format of floatingPicture to wrap square
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set textboxRange to create range boundDoc start insertionPoint end insertionPoint
set textBox1 to make new text box at boundDoc with properties {anchor:textboxRange, left position:72, top:360, width:220, height:54}
set textBox1 to shape (count shapes of boundDoc) of boundDoc
set content of text range of text frame of textBox1 to "Textbox 中文😀"
set visible of fill format of textBox1 to true
set fore color of fill format of textBox1 to {65535, 61680, 46260}
set font size of font object of text range of text frame of textBox1 to 13
set bold of font object of text range of text frame of textBox1 to true
set wrap type of wrap format of textBox1 to wrap top bottom
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set blockRange to create range boundDoc start insertionPoint end insertionPoint
set blockPicture to make new inline picture at blockRange with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-2zk_5x3f/probe-image.png", link to file:false, save with document:true}
set blockPicture to inline picture (count inline pictures of boundDoc) of boundDoc
set lock aspect ratio of blockPicture to true
set width of blockPicture to 84
set alternative text of blockPicture to "Block image alt"
set blockTextRange to text object of blockPicture
set alignment of paragraph format of blockTextRange to align paragraph center
set keep with next of paragraph format of blockTextRange to true
set space after of paragraph format of blockTextRange to 0
set insertionPoint to (end of content of text object of boundDoc) - 1
set followingRange to create range boundDoc start insertionPoint end insertionPoint
set content of followingRange to return & "Following 中文😀 paragraph" & return
set end of nativeRows to {"probe", 0, "word_version", version as text}
set end of nativeRows to {"probe", 0, "table_count", count tables of boundDoc}
set end of nativeRows to {"probe", 0, "inline_count", count inline pictures of boundDoc}
set end of nativeRows to {"probe", 0, "shape_count", count shapes of boundDoc}
set firstTable to table 1 of boundDoc
set end of nativeRows to {"table1", 1, "rows", number of rows of firstTable}
set end of nativeRows to {"table1", 1, "columns", number of columns of firstTable}
set end of nativeRows to {"table1", 1, "auto_fit", allow auto fit of firstTable}
set end of nativeRows to {"table1", 1, "repeat_header", heading format of row 1 of firstTable}
set end of nativeRows to {"table1", 1, "row_split", allow break across pages of row 2 of firstTable}
set firstCell to get cell from table firstTable row 1 column 1
set end of nativeRows to {"table1", 1, "first_text", content of text object of firstCell as text}
set end of nativeRows to {"table1", 1, "first_indent", first line indent of paragraph format of text object of firstCell}
set end of nativeRows to {"table1", 1, "left_indent", paragraph format left indent of paragraph format of text object of firstCell}
set end of nativeRows to {"table1", 1, "right_indent", paragraph format right indent of paragraph format of text object of firstCell}
set end of nativeRows to {"table1", 1, "space_before", space before of paragraph format of text object of firstCell}
set end of nativeRows to {"table1", 1, "space_after", space after of paragraph format of text object of firstCell}
set end of nativeRows to {"table1", 1, "vertical_center", (vertical alignment of firstCell is cell align vertical center)}
set end of nativeRows to {"table1", 1, "widths", {width of cell 1 of row 1 of firstTable, width of cell 2 of row 1 of firstTable, width of cell 3 of row 1 of firstTable}}
set end of nativeRows to {"table1", 1, "alignments", {my enumIndex(alignment of paragraph format of text object of cell 1 of row 2 of firstTable, {align paragraph left, align paragraph center, align paragraph right}), my enumIndex(alignment of paragraph format of text object of cell 2 of row 2 of firstTable, {align paragraph left, align paragraph center, align paragraph right}), my enumIndex(alignment of paragraph format of text object of cell 3 of row 2 of firstTable, {align paragraph left, align paragraph center, align paragraph right})}}
set firstTopBorder to get border firstTable which border border top
set end of nativeRows to {"table1", 1, "top_border", (line style of firstTopBorder is line style single)}
set end of nativeRows to {"table1", 1, "header_shade", background pattern color of shading of text object of firstCell}
set end of nativeRows to {"table1", 1, "header_font", color of font object of text object of firstCell}
set end of nativeRows to {"table2", 2, "rows", number of rows of table 2 of boundDoc}
set end of nativeRows to {"table2", 2, "columns", number of columns of table 2 of boundDoc}
set end of nativeRows to {"inline1", 1, "width", width of inline picture 1 of boundDoc}
set end of nativeRows to {"inline1", 1, "height", height of inline picture 1 of boundDoc}
set end of nativeRows to {"inline1", 1, "alt", alternative text of inline picture 1 of boundDoc as text}
set end of nativeRows to {"inline2", 2, "width", width of inline picture 2 of boundDoc}
set end of nativeRows to {"inline2", 2, "height", height of inline picture 2 of boundDoc}
set end of nativeRows to {"inline2", 2, "alt", alternative text of inline picture 2 of boundDoc as text}
set floatingPicture to shape 1 of boundDoc
set end of nativeRows to {"floating", 1, "type", my enumIndex(shape type of floatingPicture, {shape type picture, shape type text box})}
set end of nativeRows to {"floating", 1, "geometry", {left position of floatingPicture, top of floatingPicture, width of floatingPicture, height of floatingPicture}}
set end of nativeRows to {"floating", 1, "wrap", my enumIndex(wrap type of wrap format of floatingPicture, {wrap square, wrap top bottom, wrap inline})}
set end of nativeRows to {"floating", 1, "alt", alternative text of floatingPicture as text}
set textBox1 to shape 2 of boundDoc
set end of nativeRows to {"textbox", 2, "type", my enumIndex(shape type of textBox1, {shape type picture, shape type text box})}
set end of nativeRows to {"textbox", 2, "text", content of text range of text frame of textBox1 as text}
set end of nativeRows to {"textbox", 2, "geometry", {left position of textBox1, top of textBox1, width of textBox1, height of textBox1}}
set end of nativeRows to {"textbox", 2, "wrap", my enumIndex(wrap type of wrap format of textBox1, {wrap square, wrap top bottom, wrap inline})}
set end of nativeRows to {"textbox", 2, "fill_visible", visible of fill format of textBox1}
set end of nativeRows to {"textbox", 2, "fill_color", fore color of fill format of textBox1}
set end of nativeRows to {"textbox", 2, "font_size", font size of font object of text range of text frame of textBox1}
set end of nativeRows to {"textbox", 2, "bold", bold of font object of text range of text frame of textBox1}
set blockPicture to inline picture 3 of boundDoc
set blockTextRange to text object of blockPicture
set end of nativeRows to {"block", 3, "center", (alignment of paragraph format of blockTextRange is align paragraph center)}
set end of nativeRows to {"block", 3, "keep_with_next", keep with next of paragraph format of blockTextRange}
set end of nativeRows to {"block", 3, "space_after", space after of paragraph format of blockTextRange}
set end of nativeRows to {"block", 3, "following_text", content of paragraph (count paragraphs of boundDoc) of boundDoc as text}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
