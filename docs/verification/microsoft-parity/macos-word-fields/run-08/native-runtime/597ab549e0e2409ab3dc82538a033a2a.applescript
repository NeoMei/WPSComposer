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
set boundDoc to document "document-4f287a12ea9045e4a614809e0b232fa1.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xq3auboh/document-4f287a12ea9045e4a614809e0b232fa1.docx" then error "WPSC_STALE_DOCUMENT"
set semanticStyle to Word style (style body text) of boundDoc
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
set content of semanticRange to "目录😀" & return
set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + 5)
set style of semanticRange to semanticStyle
reset font object of semanticRange
reset paragraph format of semanticRange
set font size of font object of semanticRange to 18
set bold of font object of semanticRange to false
set color of font object of semanticRange to {0, 0, 0}
set alignment of paragraph format of semanticRange to align paragraph center
set space after of paragraph format of semanticRange to 5
set line spacing of paragraph format of semanticRange to 18
set outline level of paragraph format of semanticRange to outline level body text
set trailingPoint to (end of content of text object of boundDoc) - 1
set trailingRange to create range boundDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set indexPoint to insertionPoint
set indexRange to create range boundDoc start indexPoint end indexPoint
create new field text range indexRange field type field toc field text "\\o \"1-3\" \\h \\z" preserve formatting true
set ownField to missing value
repeat with fi from 1 to count fields of boundDoc
set candidateField to field fi of boundDoc
if (start of content of field code of candidateField) is indexPoint + 1 then set ownField to candidateField
end repeat
if ownField is missing value then error "WPSC_INDEX_CREATION_UNVERIFIED"
make new bookmark at boundDoc with properties {name:"WPSC_F_8fa6fcd94f3d4cb1b542e9d37460df", text object:field code of ownField}
if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"
set codeStart to start of content of field code of ownField
set resultEnd to end of content of result range of ownField
set insertionPoint to (end of content of text object of boundDoc) - 1
set terminalRange to create range boundDoc start insertionPoint end insertionPoint
insert break at terminalRange break type page break
set nativeRows to {{"index",codeStart,resultEnd,content of field code of ownField as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
