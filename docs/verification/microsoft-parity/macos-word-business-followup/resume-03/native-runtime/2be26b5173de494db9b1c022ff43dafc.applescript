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
set terminalRange to text object of last paragraph of boundDoc
set compactable to true
repeat with terminalCharacter in characters of (content of terminalRange as text)
if (terminalCharacter as text) is not in {character id 7, character id 9, character id 10, character id 11, character id 12, character id 13, character id 28, character id 29, character id 30, character id 31, character id 32, character id 133, character id 160, character id 5760, character id 8192, character id 8193, character id 8194, character id 8195, character id 8196, character id 8197, character id 8198, character id 8199, character id 8200, character id 8201, character id 8202, character id 8232, character id 8233, character id 8239, character id 8287, character id 12288} then set compactable to false
end repeat
if compactable then
set font size of font object of terminalRange to 1
set space before of paragraph format of terminalRange to 0
set space after of paragraph format of terminalRange to 0
set line spacing rule of paragraph format of terminalRange to line space exactly
set line spacing of paragraph format of terminalRange to 1
set keep together of paragraph format of terminalRange to false
set keep with next of paragraph format of terminalRange to false
end if
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
