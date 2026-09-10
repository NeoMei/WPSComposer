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
set boundDoc to document "document-61b8f9f86a09424e981ca61a3e8f1781.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-io9846pu/document-61b8f9f86a09424e981ca61a3e8f1781.docx" then error "WPSC_STALE_DOCUMENT"
set sourceStyle to Word style (style heading1) of boundDoc
set parentStyle to make new Word style at boundDoc with properties {name local:"WPSC Install Parent"}
set base style of parentStyle to style normal
set paragraph format left indent of paragraph format of parentStyle to 17
set base style of sourceStyle to parentStyle
set headingRange to create range boundDoc start 0 end 25
make new bookmark at boundDoc with properties {name:"WPSC_InstallHeading",text object:headingRange}
set ownSelection to selection of boundWindow
set selection start of ownSelection to 0
set selection end of ownSelection to 5
set nativeRows to {{"setup",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
