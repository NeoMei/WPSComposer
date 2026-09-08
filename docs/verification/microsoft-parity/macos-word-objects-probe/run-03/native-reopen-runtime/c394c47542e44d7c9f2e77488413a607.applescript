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
set boundDoc to document "document-f14d261b95264b12a2e44d1b523da648.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-4fwrbhkx/document-f14d261b95264b12a2e44d1b523da648.docx" then error "WPSC_STALE_DOCUMENT"
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
