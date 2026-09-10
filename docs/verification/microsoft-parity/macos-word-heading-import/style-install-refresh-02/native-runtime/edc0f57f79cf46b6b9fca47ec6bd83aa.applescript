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
set boundDoc to document "document-a7819db0358a4e858706c4db4fdc0de2.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-8l830wow/document-a7819db0358a4e858706c4db4fdc0de2.docx" then error "WPSC_STALE_DOCUMENT"
make new Word style at boundDoc with properties {name local:"WPSC Install Parent"}
set base style of Word style "WPSC Install Parent" of boundDoc to style normal
set paragraph format left indent of paragraph format of Word style "WPSC Install Parent" of boundDoc to 17
set sourceStyle to Word style (style heading1) of boundDoc
set parentStyle to Word style "WPSC Install Parent" of boundDoc
set normalName to name local of Word style (style normal) of boundDoc as text
set sourceName to name local of sourceStyle as text
set parentName to name local of parentStyle as text
set parentBase to name local of Word style (base style of parentStyle) of boundDoc as text
set sourceBase to name local of Word style (base style of sourceStyle) of boundDoc as text
set nativeRows to {{"parent",normalName,sourceName,parentName,parentBase,sourceBase,paragraph format left indent of paragraph format of parentStyle,paragraph format left indent of paragraph format of sourceStyle,built in of sourceStyle,built in of parentStyle}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
