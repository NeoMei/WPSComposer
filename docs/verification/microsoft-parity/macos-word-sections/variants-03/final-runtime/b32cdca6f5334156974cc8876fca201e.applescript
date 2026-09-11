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
set boundDoc to document "document-b7f963eb11f44db78911e6a5c9f62291.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-2mvtzh8a/document-b7f963eb11f44db78911e6a5c9f62291.docx" then error "WPSC_STALE_DOCUMENT"
set previousSectionCount to count sections of boundDoc
if (count sections of boundDoc) is not previousSectionCount + 0 then error "WPSC_SECTION_COUNT_FAILED"
set ownSection to section (count sections of boundDoc) of boundDoc
set inheritedOrientation to orientation of page setup of ownSection
set orientation of page setup of ownSection to orient portrait
set page width of page setup of ownSection to 612
set page height of page setup of ownSection to 1008
set orientation of page setup of ownSection to orient portrait
set top margin of page setup of ownSection to 40
set bottom margin of page setup of ownSection to 42
set left margin of page setup of ownSection to 50
set right margin of page setup of ownSection to 52
set sectionDifference to (page width of page setup of ownSection) - (612)
if sectionDifference > 0.1 or sectionDifference < -0.1 then error "WPSC_SECTION_READBACK_FAILED"
set sectionDifference to (page height of page setup of ownSection) - (1008)
if sectionDifference > 0.1 or sectionDifference < -0.1 then error "WPSC_SECTION_READBACK_FAILED"
if orientation of page setup of ownSection is not orient portrait then error "WPSC_SECTION_READBACK_FAILED"
set sectionDifference to (top margin of page setup of ownSection) - (40)
if sectionDifference > 0.1 or sectionDifference < -0.1 then error "WPSC_SECTION_READBACK_FAILED"
set sectionDifference to (bottom margin of page setup of ownSection) - (42)
if sectionDifference > 0.1 or sectionDifference < -0.1 then error "WPSC_SECTION_READBACK_FAILED"
set sectionDifference to (left margin of page setup of ownSection) - (50)
if sectionDifference > 0.1 or sectionDifference < -0.1 then error "WPSC_SECTION_READBACK_FAILED"
set sectionDifference to (right margin of page setup of ownSection) - (52)
if sectionDifference > 0.1 or sectionDifference < -0.1 then error "WPSC_SECTION_READBACK_FAILED"
set roleName to "WpsComposerSectionRole_" & ((count sections of boundDoc) as text)
set roleMatches to {}
repeat with roleIndex from 1 to (count variables of boundDoc)
set candidateRole to variable roleIndex of boundDoc
if (name of candidateRole as text) is roleName then set end of roleMatches to candidateRole
end repeat
if (count roleMatches) > 1 then error "WPSC_ROLE_AMBIGUOUS"
if (count roleMatches) is 0 then
make new variable at boundDoc with properties {name:roleName, variable value:"正文😀"}
end if
set ownRole to variable roleName of boundDoc
set variable value of ownRole to "正文😀"
if (variable value of ownRole as text) is not "正文😀" then error "WPSC_ROLE_READBACK_FAILED"
set ownHeader to get header ownSection index header footer primary
set ownFooter to get footer ownSection index header footer primary
set content of text object of ownHeader to "Continuation Header"
if (content of text object of ownHeader as text) is not "Continuation Header" & return then error "WPSC_HEADER_TEXT_FAILED"
set alignment of paragraph format of text object of ownHeader to align paragraph center
set pageBorder to get border (paragraph format of text object of ownHeader) which border border bottom
set line style of pageBorder to line style single
set line width of pageBorder to line width75 point
set color of pageBorder to {0, 0, 0}
set content of text object of ownFooter to "Variant "
if (content of text object of ownFooter as text) is not "Variant " & return then error "WPSC_HEADER_TEXT_FAILED"
set ownFooter to get footer ownSection index header footer primary
set footerRange to text object of ownFooter
set pageCount to 0
repeat with fieldIndex from 1 to (count fields of footerRange)
set ownField to field fieldIndex of footerRange
if field type of ownField is field page then set pageCount to pageCount + 1
end repeat
repeat with fieldIndex from (count fields of footerRange) to 1 by -1
set ownField to field fieldIndex of footerRange
if field type of ownField is field page then
if pageCount > 1 then
delete ownField
set pageCount to pageCount - 1
end if
end if
end repeat
if pageCount is 0 then
set footerRange to text object of ownFooter
set lastCharacter to count characters of footerRange
create new field text range (character lastCharacter of text object of ownFooter) field type field page preserve formatting true
end if
set alignment of paragraph format of text object of ownFooter to align paragraph center
set number style of page number options of ownFooter to page number style arabic
set restart numbering at section of page number options of ownFooter to true
set starting number of page number options of ownFooter to 9
if number style of page number options of ownFooter is not page number style arabic then error "WPSC_NUMBERING_READBACK_FAILED"
if restart numbering at section of page number options of ownFooter is not true then error "WPSC_NUMBERING_READBACK_FAILED"
if starting number of page number options of ownFooter is not 9 then error "WPSC_NUMBERING_READBACK_FAILED"
set footerRange to text object of ownFooter
set pageCount to 0
repeat with fieldIndex from 1 to (count fields of footerRange)
set ownField to field fieldIndex of footerRange
if field type of ownField is field page then
if (word 1 of (content of field code of ownField as text) as text) is not "PAGE" then error "WPSC_PAGE_COMMAND_INVALID"
set pageCount to pageCount + 1
end if
end repeat
if pageCount is not 1 then error "WPSC_PAGE_COUNT_FAILED"
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
