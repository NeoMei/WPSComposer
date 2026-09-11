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
set boundDoc to document "document-630ad4e46b324d08869b4fab94ea7331.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-bu3wg936/document-630ad4e46b324d08869b4fab94ea7331.docx" then error "WPSC_STALE_DOCUMENT"
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
delete ownField
set pageCount to pageCount - 1
end if
end repeat
set footerRange to text object of ownFooter
set pageCount to 0
repeat with fieldIndex from 1 to (count fields of footerRange)
set ownField to field fieldIndex of footerRange
if field type of ownField is field page then
if (word 1 of (content of field code of ownField as text) as text) is not "PAGE" then error "WPSC_PAGE_COMMAND_INVALID"
set pageCount to pageCount + 1
end if
end repeat
if pageCount is not 0 then error "WPSC_PAGE_COUNT_FAILED"
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
