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
set boundDoc to document "document-3681b4e8d0f04137a279e49b159ecf47.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5ylgul1n/document-3681b4e8d0f04137a279e49b159ecf47.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {{"counts", count tables of boundDoc, count inline pictures of boundDoc, count shapes of boundDoc}}
set end of nativeRows to {"table", number of rows of table 1 of boundDoc, number of columns of table 1 of boundDoc, heading format of row 1 of table 1 of boundDoc, allow break across pages of row 2 of table 1 of boundDoc}
set c to get cell from table (table 1 of boundDoc) row 1 column 1
set end of nativeRows to {"cell", content of text object of c as text, first line indent of paragraph format of text object of c, paragraph format left indent of paragraph format of text object of c, paragraph format right indent of paragraph format of text object of c, vertical alignment of c is cell align vertical center}
set end of nativeRows to {"inline1", width of inline picture 1 of boundDoc, height of inline picture 1 of boundDoc, alternative text of inline picture 1 of boundDoc as text}
set end of nativeRows to {"inline2", width of inline picture 2 of boundDoc, height of inline picture 2 of boundDoc}
set end of nativeRows to {"inline3", width of inline picture 3 of boundDoc, height of inline picture 3 of boundDoc}
set end of nativeRows to {"inline4", width of inline picture 4 of boundDoc, height of inline picture 4 of boundDoc}
set blockRange to text object of inline picture 5 of boundDoc
set end of nativeRows to {"block", alignment of paragraph format of blockRange is align paragraph center, keep with next of paragraph format of blockRange, space after of paragraph format of blockRange}
set floatingPicture to shape 1 of boundDoc
set end of nativeRows to {"floating", shape type of floatingPicture is shape type picture, width of floatingPicture, height of floatingPicture, wrap type of wrap format of floatingPicture is wrap square}
set textBoxShape to shape 2 of boundDoc
set end of nativeRows to {"textbox", shape type of textBoxShape is shape type text box, content of text range of text frame of textBoxShape as text, left position of textBoxShape, top of textBoxShape, width of textBoxShape, height of textBoxShape, wrap type of wrap format of textBoxShape is wrap square, font size of font object of text range of text frame of textBoxShape, bold of font object of text range of text frame of textBoxShape}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
