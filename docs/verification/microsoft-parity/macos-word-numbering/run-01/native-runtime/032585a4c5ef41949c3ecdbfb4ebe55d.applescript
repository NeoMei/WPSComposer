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
activate object boundWindow
set content of text object of boundDoc to "Chapter" & return & "left REPLACE right" & return & "TAIL" & return
set numberList to make new list template at boundDoc with properties {name:"WPSCNumberingFixture",outline numbered:true}
set numberLevel to list level 1 of numberList
set number style of numberLevel to list number style arabic
set number format of numberLevel to "%1"
set start at of numberLevel to 1
set linked style of numberLevel to name local of Word style (style heading1) of boundDoc as text
set style of text object of paragraph 1 of boundDoc to style heading1
set selection start of selection of boundWindow to 13
set selection end of selection of boundWindow to 20
set nativeRows to {{"seed",list string of list format of text object of paragraph 1 of boundDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
