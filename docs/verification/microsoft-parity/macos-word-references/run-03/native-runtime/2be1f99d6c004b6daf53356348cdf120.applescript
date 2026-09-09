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
set boundDoc to document "document-3e8ba73679464786b4aa35ba0a39180e.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-f0jk6515/document-3e8ba73679464786b4aa35ba0a39180e.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
if insertionPoint > 0 then
set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint
if (content of precedingRange as text) is not return then
set boundaryRange to create range boundDoc start insertionPoint end insertionPoint
set content of boundaryRange to return
set insertionPoint to insertionPoint + 1
end if
end if
set operationStart to insertionPoint
set beforeParagraphs to count paragraphs of boundDoc
set nativeRows to {}
set referenceFailed to false
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "1." & tab & ""
set insertionPoint to insertionPoint + 3
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "1." & tab & "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",0,literalStart,insertionPoint,content of literalRange as text}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "[3]"
set insertionPoint to insertionPoint + 3
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "[3]" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "[MISS 缺失]"
set insertionPoint to insertionPoint + 9
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "[MISS 缺失]" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",2,literalStart,insertionPoint,content of literalRange as text}
set italic of font object of literalRange to true
set color of font object of literalRange to {40092, 0, 1542}
set background pattern color of shading of literalRange to {64764, 59624, 59110}
set end of nativeRows to {"styled",2,literalStart,insertionPoint,italic of font object of literalRange,(color of font object of literalRange is {40092, 0, 1542}),(background pattern color of shading of literalRange is {64764, 59624, 59110})}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "[MISS 缺失]"
set insertionPoint to insertionPoint + 9
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "[MISS 缺失]" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",3,literalStart,insertionPoint,content of literalRange as text}
set italic of font object of literalRange to true
set color of font object of literalRange to {40092, 0, 1542}
set background pattern color of shading of literalRange to {64764, 59624, 59110}
set end of nativeRows to {"styled",3,literalStart,insertionPoint,italic of font object of literalRange,(color of font object of literalRange is {40092, 0, 1542}),(background pattern color of shading of literalRange is {64764, 59624, 59110})}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "" & return & ""
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "" & return & "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",4,literalStart,insertionPoint,content of literalRange as text}
set paragraphRange to create range boundDoc start operationStart end insertionPoint
try
set style of paragraphRange to Word style "List Paragraph" of boundDoc
end try
set paragraph format left indent of paragraph format of paragraphRange to 31.5
set first line indent of paragraph format of paragraphRange to -31.5
make new tab stop at paragraph 1 of paragraphRange with properties {tab stop position:31.5}
set line spacing rule of paragraph format of paragraphRange to line space1 pt5
set space before of paragraph format of paragraphRange to 0
set space after of paragraph format of paragraphRange to 3
set tabVerified to false
repeat with tabIndex from 1 to count tab stops of paragraph 1 of paragraphRange
set ownTab to tab stop tabIndex of paragraph 1 of paragraphRange
if (tab stop position of ownTab) is 31.5 then set tabVerified to true
end repeat
set end of nativeRows to {"list",paragraph format left indent of paragraph format of paragraphRange,first line indent of paragraph format of paragraphRange,(line spacing rule of paragraph format of paragraphRange is line space1 pt5),space before of paragraph format of paragraphRange,space after of paragraph format of paragraphRange,tabVerified}
set trailingRange to create range boundDoc start insertionPoint end insertionPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set end of nativeRows to {"complete",beforeParagraphs,count paragraphs of boundDoc,operationStart,insertionPoint}
end if
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
