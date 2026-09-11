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
set boundDoc to document "document-928d27e1cc144271a6a764b030886d82.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-fk90_aps/document-928d27e1cc144271a6a764b030886d82.docx" then error "WPSC_STALE_DOCUMENT"
set semanticSection to section (count sections of boundDoc) of boundDoc
if (count sections of boundDoc) > 1 then
set pagePart to get header semanticSection index header footer primary
set link to previous of pagePart to false
if link to previous of pagePart is not false then error "WPSC_HEADER_LINK_FAILED"
end if
if (count sections of boundDoc) > 1 then
set pagePart to get footer semanticSection index header footer primary
set link to previous of pagePart to true
if link to previous of pagePart is not true then error "WPSC_HEADER_LINK_FAILED"
end if
set pagePart to get header semanticSection index header footer primary
set content of text object of pagePart to "Section Two"
set alignment of paragraph format of text object of pagePart to align paragraph center
set pageBorder to get border (paragraph format of text object of pagePart) which border border bottom
set line style of pageBorder to line style single
set line width of pageBorder to line width75 point
set color of pageBorder to {0, 0, 0}
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
