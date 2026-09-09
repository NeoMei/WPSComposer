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
set boundDoc to document "document-651fc661d3a64ce6a966ffc90e2e5750.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-hhk9noow/document-651fc661d3a64ce6a966ffc90e2e5750.docx" then error "WPSC_STALE_DOCUMENT"
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
set styleRanges to {}
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set emptyValue to content of literalRange
if emptyValue is not missing value then
if (emptyValue as text) is not "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
end if
set end of nativeRows to {"literal",0,start of content of literalRange,end of content of literalRange,""}
set referenceDegraded to false
set referenceStart to insertionPoint
set previousDocumentEnd to end of content of text object of boundDoc
set previousFields to count fields of boundDoc
try
set ownRange to create range boundDoc start referenceStart end referenceStart
create new field text range ownRange field type field ref field text "wpsc_eq_aaaaaaaaaaaaaaaaaaaaaaaa \\h" preserve formatting true
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
if referenceCode is not " REF wpsc_eq_aaaaaaaaaaaaaaaaaaaaaaaa \\h " and referenceCode is not " REF wpsc_eq_aaaaaaaaaaaaaaaaaaaaaaaa \\h \\* MERGEFORMAT " and referenceCode is not "REF wpsc_eq_aaaaaaaaaaaaaaaaaaaaaaaa \\h" then error "WPSC_REFERENCE_CODE_UNVERIFIED"
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
set end of nativeRows to {"fallback",0,referenceStart,actualPartialEnd,previousFields,count fields of boundDoc}
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "BAD-REF"
set insertionPoint to insertionPoint + 7
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "BAD-REF" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",0,literalStart,insertionPoint,content of literalRange as text}
end try
if not referenceFailed and not referenceDegraded then
make new bookmark at boundDoc with properties {name:"WPSC_R_e3250ec6fa444ec2960bfda70a6a0f", text object:field code of ownField}
if (content of text object of bookmark "WPSC_R_e3250ec6fa444ec2960bfda70a6a0f" of boundDoc as text) is not referenceCode then error "WPSC_REFERENCE_IDENTITY_UNVERIFIED"
set end of nativeRows to {"reference",0,previousFields,count fields of boundDoc,referenceStart + 1,referenceResultStart,referenceResultEnd,referenceCode,"WPSC_R_e3250ec6fa444ec2960bfda70a6a0f"}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to ";"
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not ";" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",0,literalStart,insertionPoint,content of literalRange as text}
end if
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
create new field text range ownRange field type field ref field text "wpsc_eq_bbbbbbbbbbbbbbbbbbbbbbbb \\h" preserve formatting true
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
if referenceCode is not " REF wpsc_eq_bbbbbbbbbbbbbbbbbbbbbbbb \\h " and referenceCode is not " REF wpsc_eq_bbbbbbbbbbbbbbbbbbbbbbbb \\h \\* MERGEFORMAT " and referenceCode is not "REF wpsc_eq_bbbbbbbbbbbbbbbbbbbbbbbb \\h" then error "WPSC_REFERENCE_CODE_UNVERIFIED"
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
set content of literalRange to "BAD-REF"
set insertionPoint to insertionPoint + 7
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "BAD-REF" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
end try
if not referenceFailed and not referenceDegraded then
make new bookmark at boundDoc with properties {name:"WPSC_R_b795c73481134233bb2f4832cabe25", text object:field code of ownField}
if (content of text object of bookmark "WPSC_R_b795c73481134233bb2f4832cabe25" of boundDoc as text) is not referenceCode then error "WPSC_REFERENCE_IDENTITY_UNVERIFIED"
set end of nativeRows to {"reference",1,previousFields,count fields of boundDoc,referenceStart + 1,referenceResultStart,referenceResultEnd,referenceCode,"WPSC_R_b795c73481134233bb2f4832cabe25"}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to ";"
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not ";" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",1,literalStart,insertionPoint,content of literalRange as text}
end if
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set emptyValue to content of literalRange
if emptyValue is not missing value then
if (emptyValue as text) is not "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
end if
set end of nativeRows to {"literal",2,start of content of literalRange,end of content of literalRange,""}
set referenceDegraded to false
set referenceStart to insertionPoint
set previousDocumentEnd to end of content of text object of boundDoc
set previousFields to count fields of boundDoc
try
set ownRange to create range boundDoc start referenceStart end referenceStart
create new field text range ownRange field type field ref field text "wpsc_eq_cccccccccccccccccccccccc \\h" preserve formatting true
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
if referenceCode is not " REF wpsc_eq_cccccccccccccccccccccccc \\h " and referenceCode is not " REF wpsc_eq_cccccccccccccccccccccccc \\h \\* MERGEFORMAT " and referenceCode is not "REF wpsc_eq_cccccccccccccccccccccccc \\h" then error "WPSC_REFERENCE_CODE_UNVERIFIED"
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
set end of nativeRows to {"fallback",2,referenceStart,actualPartialEnd,previousFields,count fields of boundDoc}
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "BAD-REF"
set insertionPoint to insertionPoint + 7
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "BAD-REF" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",2,literalStart,insertionPoint,content of literalRange as text}
end try
if not referenceFailed and not referenceDegraded then
make new bookmark at boundDoc with properties {name:"WPSC_R_afaa0c6458f1498db6a84fbccbf861", text object:field code of ownField}
if (content of text object of bookmark "WPSC_R_afaa0c6458f1498db6a84fbccbf861" of boundDoc as text) is not referenceCode then error "WPSC_REFERENCE_IDENTITY_UNVERIFIED"
set end of nativeRows to {"reference",2,previousFields,count fields of boundDoc,referenceStart + 1,referenceResultStart,referenceResultEnd,referenceCode,"WPSC_R_afaa0c6458f1498db6a84fbccbf861"}
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to ";"
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not ";" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",2,literalStart,insertionPoint,content of literalRange as text}
end if
end if
if not referenceFailed then
set literalStart to insertionPoint
set literalRange to create range boundDoc start insertionPoint end insertionPoint
set content of literalRange to "" & return & ""
set insertionPoint to insertionPoint + 1
set literalRange to create range boundDoc start literalStart end insertionPoint
if (content of literalRange as text) is not "" & return & "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"
set end of nativeRows to {"literal",3,literalStart,insertionPoint,content of literalRange as text}
repeat with styleIndex from 1 to count styleRanges
set styleSpec to item styleIndex of styleRanges
set styledRange to create range boundDoc start (item 2 of styleSpec) end (item 3 of styleSpec)
set italic of font object of styledRange to true
set color of font object of styledRange to {40092, 0, 1542}
set background pattern color of shading of styledRange to {64764, 59624, 59110}
set end of nativeRows to {"styled",item 1 of styleSpec,start of content of styledRange,end of content of styledRange,italic of font object of styledRange,(color of font object of styledRange is {40092, 0, 1542}),(background pattern color of shading of styledRange is {64764, 59624, 59110})}
end repeat
set end of nativeRows to {"complete",beforeParagraphs,count paragraphs of boundDoc,operationStart,insertionPoint}
end if
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
