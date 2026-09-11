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
set boundDoc to document "document-e4174e18b98b43ccb46eb85bf5d874b2.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mx2350s5/document-e4174e18b98b43ccb46eb85bf5d874b2.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mx2350s5/document-e4174e18b98b43ccb46eb85bf5d874b2.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mx2350s5/document-e4174e18b98b43ccb46eb85bf5d874b2.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mx2350s5/document-e4174e18b98b43ccb46eb85bf5d874b2.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityStoryRows to {}
set end of qualityStoryRows to {"story-block","first-header"}
set qualityStoryRange to missing value
try
set qualityStoryRange to get story range boundDoc story type first page header story
set qualityRootClass to (class of qualityStoryRange) as text
if qualityRootClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",0,"root","empty-class"}
else if qualityStoryRange is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",0,"root","native-missing"}
else
set qualityStoryNode to 1
repeat
set end of qualityStoryRows to {"story-node","first-header",qualityStoryNode as integer}
set probePropertyName to "class"
set probeValue to missing value
try
set probeValue to get (class of qualityStoryRange) as text
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"property-ok","first-header",qualityStoryNode,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "story-type"
set probeValue to missing value
try
set probeValue to get story type of qualityStoryRange
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"property-ok","first-header",qualityStoryNode,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "start"
set probeValue to missing value
try
set probeValue to get start of content of qualityStoryRange
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"property-ok","first-header",qualityStoryNode,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "end"
set probeValue to missing value
try
set probeValue to get end of content of qualityStoryRange
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"property-ok","first-header",qualityStoryNode,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "content"
set probeValue to missing value
try
set probeValue to get content of qualityStoryRange
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"property-ok","first-header",qualityStoryNode,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "section-count"
set probeValue to missing value
try
set probeValue to get (count sections of qualityStoryRange) as integer
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"property-ok","first-header",qualityStoryNode,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "section-first-index"
set probeValue to missing value
try
set probeValue to get section index of section 1 of qualityStoryRange
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"property-ok","first-header",qualityStoryNode,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set qualityFieldCountState to "error"
set qualityStoryFieldCount to 0
try
set qualityStoryFieldCount to (count fields of qualityStoryRange) as integer
set qualityFieldCountState to "ok"
on error qualityFieldCountError number qualityFieldCountNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,"field-count",qualityFieldCountNumber as integer,qualityFieldCountError as text}
end try
if qualityFieldCountState is "ok" then
set end of qualityStoryRows to {"field-count","first-header",qualityStoryNode,qualityStoryFieldCount}
repeat with qualityFieldOrdinal from 1 to qualityStoryFieldCount
set qualityStoryField to field qualityFieldOrdinal of qualityStoryRange
set end of qualityStoryRows to {"field-node","first-header",qualityStoryNode,qualityFieldOrdinal as integer}
set probePropertyName to "field-type"
set probeValue to missing value
try
set probeValue to get field type of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "field-code-content"
set probeValue to missing value
try
set probeValue to get content of field code of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "field-code-start"
set probeValue to missing value
try
set probeValue to get start of content of field code of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "field-code-end"
set probeValue to missing value
try
set probeValue to get end of content of field code of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "field-result-content"
set probeValue to missing value
try
set probeValue to get content of result range of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "field-result-start"
set probeValue to missing value
try
set probeValue to get start of content of result range of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "field-result-end"
set probeValue to missing value
try
set probeValue to get end of content of result range of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
set probePropertyName to "field-locked"
set probeValue to missing value
try
set probeValue to get locked of qualityStoryField
set probeValueClass to (class of probeValue) as text
if probeValueClass is "" then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"empty-class"}
else if probeValue is missing value then
set end of qualityStoryRows to {"unresolved-field-property","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,"native-missing"}
else
if probeValueClass is "boolean" or probeValueClass is "integer" or probeValueClass is "real" or probeValueClass is "text" then
set probeSerializableValue to probeValue
else
set probeSerializableValue to probeValue as text
end if
set end of qualityStoryRows to {"field-property-ok","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probeValueClass,probeSerializableValue}
end if
on error probePropertyError number probePropertyNumber
set end of qualityStoryRows to {"field-property-error","first-header",qualityStoryNode,qualityFieldOrdinal,probePropertyName,probePropertyNumber as integer,probePropertyError as text}
end try
end repeat
end if
set qualityNextStoryRange to missing value
try
set qualityNextStoryRange to get next story range of qualityStoryRange
set qualityNextClass to (class of qualityNextStoryRange) as text
if qualityNextClass is "" then
set end of qualityStoryRows to {"unresolved-property","first-header",qualityStoryNode,"next-story","empty-class"}
exit repeat
else if qualityNextStoryRange is missing value then
set end of qualityStoryRows to {"chain-end","first-header",qualityStoryNode,"missing-value"}
exit repeat
else
set qualityStoryRange to qualityNextStoryRange
set qualityStoryNode to qualityStoryNode + 1
if qualityStoryNode > 64 then error "WPSC_STORY_CHAIN_BOUND_EXCEEDED"
end if
on error qualityNextError number qualityNextNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",qualityStoryNode,"next-story",qualityNextNumber as integer,qualityNextError as text}
exit repeat
end try
end repeat
end if
on error qualityRootError number qualityRootNumber
set end of qualityStoryRows to {"ordinary-property-error","first-header",0,"root",qualityRootNumber as integer,qualityRootError as text}
end try
set nativeRows to {{"story-snapshot",qualityStoryRows}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
