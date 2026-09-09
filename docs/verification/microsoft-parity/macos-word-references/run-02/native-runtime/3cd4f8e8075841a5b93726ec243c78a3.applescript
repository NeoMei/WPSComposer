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
set boundDoc to document "document-6fe260f50f6e4df181b7fcbc16f27abc.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-eddiheun/document-6fe260f50f6e4df181b7fcbc16f27abc.docx" then error "WPSC_STALE_DOCUMENT"
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
set content of literalRange to "•" & tab & ""
set insertionPoint to insertionPoint + 2
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "•" & tab & "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",0,literalStart,insertionPoint,content of literalRange as text}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set emptyValue to content of literalRange
if emptyValue is not missing value then
if (emptyValue as text) is not "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
end if
set end of nativeRows to {"literal",1,start of content of literalRange,end of content of literalRange,""}
set referenceDegraded to false
set referenceStart to insertionPoint
set previousDocumentEnd to end of content of text object of boundDoc
set previousFields to count fields of boundDoc
try
set ownRange to create range boundDoc start referenceStart end referenceStart
create new field text range ownRange field type field ref field text "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h" preserve formatting true
if (count fields of boundDoc) is not previousFields + 1 then error "WPSC_REFERENCE_COUNT_UNVERIFIED"
set ownField to missing value
set matchingFields to 0
repeat with fi from 1 to count fields of boundDoc
set candidateField to field fi of boundDoc
if (start of content of field code of candidateField) is referenceStart + 1 then
set ownField to candidateField
set matchingFields to matchingFields + 1
end if
end repeat
if matchingFields is not 1 then error "WPSC_REFERENCE_IDENTITY_UNVERIFIED"
if field type of ownField is not field ref then error "WPSC_REFERENCE_TYPE_UNVERIFIED"
set referenceCode to content of field code of ownField as text
if referenceCode is not " REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h " and referenceCode is not " REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT " and referenceCode is not "REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\h" then error "WPSC_REFERENCE_CODE_UNVERIFIED"
set referenceResultStart to start of content of result range of ownField
set referenceResultEnd to end of content of result range of ownField
if referenceResultStart < referenceStart + 2 or referenceResultEnd < referenceResultStart then error "WPSC_REFERENCE_RANGE_UNVERIFIED"
set insertionPoint to referenceResultEnd + 1
if insertionPoint > (end of content of text object of boundDoc) - 1 then error "WPSC_REFERENCE_RANGE_UNVERIFIED"
on error referenceError number referenceNumber
if referenceError starts with "WPSC_REFERENCE_" then error referenceError number referenceNumber
set actualPartialEnd to (end of content of text object of boundDoc) - 1
if actualPartialEnd < referenceStart then error "WPSC_REFERENCE_ROLLBACK_UNVERIFIED"
set rollbackRange to create range boundDoc start referenceStart end actualPartialEnd
set content of rollbackRange to ""
if (end of content of text object of boundDoc) is not previousDocumentEnd or (count fields of boundDoc) is not previousFields then error "WPSC_REFERENCE_ROLLBACK_UNVERIFIED"
set insertionPoint to referenceStart
set referenceDegraded to true
set end of nativeRows to {"fallback",1,referenceStart,actualPartialEnd,previousFields,count fields of boundDoc}
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "7"
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "7" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
end try
if not referenceFailed and not referenceDegraded then
make new bookmark at boundDoc with properties {name:"WPSC_R_6eafb16bfe1f45b08d1ebe95888e54", text object:field code of ownField}
if (content of text object of bookmark "WPSC_R_6eafb16bfe1f45b08d1ebe95888e54" of boundDoc as text) is not referenceCode then error "WPSC_REFERENCE_IDENTITY_UNVERIFIED"
set end of nativeRows to {"reference",1,previousFields,count fields of boundDoc,referenceStart + 1,referenceResultStart,referenceResultEnd,referenceCode,"WPSC_R_6eafb16bfe1f45b08d1ebe95888e54"}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set emptyValue to content of literalRange
if emptyValue is not missing value then
if (emptyValue as text) is not "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
end if
set end of nativeRows to {"literal",1,start of content of literalRange,end of content of literalRange,""}
end if
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "" & return & ""
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "" & return & "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",2,literalStart,insertionPoint,content of literalRange as text}
set paragraphRange to create range boundDoc start operationStart end insertionPoint
try
set style of paragraphRange to Word style "List Paragraph" of boundDoc
end try
set paragraph format left indent of paragraph format of paragraphRange to 24.0
set first line indent of paragraph format of paragraphRange to -24.0
make new tab stop at paragraph 1 of paragraphRange with properties {tab stop position:24.0}
set line spacing rule of paragraph format of paragraphRange to line space1 pt5
set space before of paragraph format of paragraphRange to 0
set space after of paragraph format of paragraphRange to 3
set tabVerified to false
repeat with ownTab in tab stops of paragraph 1 of paragraphRange
if (tab stop position of ownTab) is 24.0 then set tabVerified to true
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
