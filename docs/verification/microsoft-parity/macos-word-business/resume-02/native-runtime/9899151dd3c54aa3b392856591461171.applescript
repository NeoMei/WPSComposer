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
set boundDoc to document "document-3492017ce5a4493299ca1450fc30b6b9.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-oqpusb84/document-3492017ce5a4493299ca1450fc30b6b9.docx" then error "WPSC_STALE_DOCUMENT"
set semanticStyle to Word style (style list paragraph) of boundDoc
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set semanticRange to create range boundDoc start insertionPoint end insertionPoint
set content of semanticRange to "1." & tab & "First literal number" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 24)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set paragraph format left indent of paragraph format of semanticRange to 24
set first line indent of paragraph format of semanticRange to -24
set space before of paragraph format of semanticRange to 0
set space after of paragraph format of semanticRange to 3
set font size of font object of semanticRange to 12
set bold of font object of semanticRange to false
set color of font object of semanticRange to {0, 0, 0}
set name of font object of semanticRange to "仿宋"
set east asian name of font object of semanticRange to "仿宋"
set ascii name of font object of semanticRange to "Times New Roman"
set other name of font object of semanticRange to "Times New Roman"
set complex script name of font object of semanticRange to "Times New Roman"
set line spacing rule of paragraph format of semanticRange to line space1 pt5
make new tab stop at paragraph 1 of semanticRange with properties {tab stop position:24}
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set semanticStyle to Word style (style list paragraph) of boundDoc
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set semanticRange to create range boundDoc start insertionPoint end insertionPoint
set content of semanticRange to "2." & tab & "Second literal number" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 25)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set paragraph format left indent of paragraph format of semanticRange to 24
set first line indent of paragraph format of semanticRange to -24
set space before of paragraph format of semanticRange to 0
set space after of paragraph format of semanticRange to 3
set font size of font object of semanticRange to 12
set bold of font object of semanticRange to false
set color of font object of semanticRange to {0, 0, 0}
set name of font object of semanticRange to "仿宋"
set east asian name of font object of semanticRange to "仿宋"
set ascii name of font object of semanticRange to "Times New Roman"
set other name of font object of semanticRange to "Times New Roman"
set complex script name of font object of semanticRange to "Times New Roman"
set line spacing rule of paragraph format of semanticRange to line space1 pt5
make new tab stop at paragraph 1 of semanticRange with properties {tab stop position:24}
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
