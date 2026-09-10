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
set boundDoc to document "document-3afcfe178d8a46b3866aa6c42d1408b4.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5u76h90z/document-3afcfe178d8a46b3866aa6c42d1408b4.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5u76h90z/document-3afcfe178d8a46b3866aa6c42d1408b4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5u76h90z/document-3afcfe178d8a46b3866aa6c42d1408b4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5u76h90z/document-3afcfe178d8a46b3866aa6c42d1408b4.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityStyleNames to get name local of every Word style of boundDoc
set qualityStyleInUseFlags to get in use of every Word style of boundDoc
set qualityStyleBuiltinFlags to get built in of every Word style of boundDoc
if (count qualityStyleNames) is not (count qualityStyleInUseFlags) or (count qualityStyleNames) is not (count qualityStyleBuiltinFlags) then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set qualityStyleRows to {{"style-catalog",(count qualityStyleNames) as integer}}
repeat with qualityStyleOrdinal from 1 to count qualityStyleNames
set qualityStyleCurrentName to item qualityStyleOrdinal of qualityStyleNames as text
set qualityStyleIsInUse to item qualityStyleOrdinal of qualityStyleInUseFlags
set qualityStyleIsBuiltin to item qualityStyleOrdinal of qualityStyleBuiltinFlags
if class of qualityStyleIsInUse is not boolean or class of qualityStyleIsBuiltin is not boolean then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set end of qualityStyleRows to {"style-state",qualityStyleOrdinal as integer,qualityStyleCurrentName,qualityStyleIsInUse,qualityStyleIsBuiltin}
if qualityStyleIsInUse or not qualityStyleIsBuiltin then
set qualityCurrentStyle to Word style qualityStyleOrdinal of boundDoc
set qualityResolvedStyleName to name local of qualityCurrentStyle as text
if not ((current application's NSString's stringWithString:qualityResolvedStyleName)'s isEqualToString:qualityStyleCurrentName) then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set qualityStyleDescription to description of qualityCurrentStyle as text
set qualityStyleAutomaticallyUpdates to automatically update of qualityCurrentStyle
set end of qualityStyleRows to {"style-definition",qualityStyleOrdinal as integer,qualityResolvedStyleName,qualityStyleDescription,qualityStyleAutomaticallyUpdates}
end if
end repeat
set qualityStyleNamesAfter to get name local of every Word style of boundDoc
if not ((current application's NSString's stringWithString:(my jsonRows({qualityStyleNamesAfter})))'s isEqualToString:(my jsonRows({qualityStyleNames}))) then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"
set nativeRows to {{"style-snapshot",qualityStyleRows}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
