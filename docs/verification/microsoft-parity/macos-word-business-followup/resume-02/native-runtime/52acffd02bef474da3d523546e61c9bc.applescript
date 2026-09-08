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
set boundDoc to document "document-14e9eb489ac54811b63255c8e4507391.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-1ylga8vo/document-14e9eb489ac54811b63255c8e4507391.docx" then error "WPSC_STALE_DOCUMENT"
set requestedStyle0 to missing value
try
set requestedStyle0 to Word style "Source Code" of boundDoc
end try
if requestedStyle0 is not missing value then
if (style type of requestedStyle0) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
end if
set requestedStyle1 to missing value
try
set requestedStyle1 to Word style "Custom Body" of boundDoc
end try
if requestedStyle1 is not missing value then
if (style type of requestedStyle1) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
end if
set baseStyle1 to Word style (style body text) of boundDoc
if (style type of baseStyle1) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
set semanticStyle to requestedStyle0
if semanticStyle is missing value then
set semanticStyle to make new Word style at boundDoc with properties {name local:"Source Code"}
end if
set paragraph format left indent of paragraph format of semanticStyle to 36
set first line indent of paragraph format of semanticStyle to 0
set space after of paragraph format of semanticStyle to 0
set space before of paragraph format of semanticStyle to 0
set font size of font object of semanticStyle to 9
set name of font object of semanticStyle to "Consolas"
set ascii name of font object of semanticStyle to "Consolas"
set other name of font object of semanticStyle to "Consolas"
set complex script name of font object of semanticStyle to "Consolas"
set line spacing rule of paragraph format of semanticStyle to line space single
set background pattern color of shading of semanticStyle to {62965, 62965, 62965}
set requestedStyle0 to semanticStyle
set semanticStyle to requestedStyle1
if semanticStyle is missing value then
set semanticStyle to make new Word style at boundDoc with properties {name local:"Custom Body"}
end if
set base style of semanticStyle to baseStyle1
set color of font object of semanticStyle to {4626, 13364, 22102}
set first line indent of paragraph format of semanticStyle to 24
set space after of paragraph format of semanticStyle to 4
set font size of font object of semanticStyle to 12
set requestedStyle1 to semanticStyle
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
