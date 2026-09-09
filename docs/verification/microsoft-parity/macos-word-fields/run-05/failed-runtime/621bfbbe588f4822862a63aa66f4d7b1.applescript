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
set boundDoc to document "document-de797c88dbce4c9c90073173f3a023ec.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-dvgjer62/document-de797c88dbce4c9c90073173f3a023ec.docx" then error "WPSC_STALE_DOCUMENT"
set ownSection to section (count sections of boundDoc) of boundDoc
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
set starting number of page number options of ownFooter to 1
if number style of page number options of ownFooter is not page number style arabic then error "WPSC_NUMBERING_READBACK_FAILED"
if restart numbering at section of page number options of ownFooter is not true then error "WPSC_NUMBERING_READBACK_FAILED"
if starting number of page number options of ownFooter is not 1 then error "WPSC_NUMBERING_READBACK_FAILED"
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
