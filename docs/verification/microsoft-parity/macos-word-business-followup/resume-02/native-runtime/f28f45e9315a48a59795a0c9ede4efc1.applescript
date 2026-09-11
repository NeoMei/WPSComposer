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
set requestedHeading1 to Word style (style heading1) of boundDoc
if (style type of requestedHeading1) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
set requestedHeading2 to Word style (style heading2) of boundDoc
if (style type of requestedHeading2) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
set requestedHeading3 to Word style (style heading3) of boundDoc
if (style type of requestedHeading3) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
set requestedHeading4 to Word style (style heading4) of boundDoc
if (style type of requestedHeading4) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
set requestedHeading5 to Word style (style heading5) of boundDoc
if (style type of requestedHeading5) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
set requestedHeading6 to Word style (style heading6) of boundDoc
if (style type of requestedHeading6) is not in {style type paragraph, style type paragraph only, style type linked} then error "WPSC_STYLE_TYPE_MISMATCH"
set semanticStyle to requestedHeading1
set bold of font object of semanticStyle to true
set alignment of paragraph format of semanticStyle to align paragraph center
set space before of paragraph format of semanticStyle to 16
set space after of paragraph format of semanticStyle to 5
set keep with next of paragraph format of semanticStyle to true
set font size of font object of semanticStyle to 16
set name of font object of semanticStyle to "黑体"
set east asian name of font object of semanticStyle to "黑体"
set ascii name of font object of semanticStyle to "黑体"
set other name of font object of semanticStyle to "黑体"
set complex script name of font object of semanticStyle to "黑体"
set line spacing rule of paragraph format of semanticStyle to line space single
set semanticStyle to requestedHeading2
set bold of font object of semanticStyle to true
set alignment of paragraph format of semanticStyle to align paragraph left
set space before of paragraph format of semanticStyle to 14
set space after of paragraph format of semanticStyle to 5
set keep with next of paragraph format of semanticStyle to true
set font size of font object of semanticStyle to 15
set name of font object of semanticStyle to "黑体"
set east asian name of font object of semanticStyle to "黑体"
set ascii name of font object of semanticStyle to "黑体"
set other name of font object of semanticStyle to "黑体"
set complex script name of font object of semanticStyle to "黑体"
set line spacing rule of paragraph format of semanticStyle to line space single
set semanticStyle to requestedHeading3
set bold of font object of semanticStyle to true
set alignment of paragraph format of semanticStyle to align paragraph left
set space before of paragraph format of semanticStyle to 12
set space after of paragraph format of semanticStyle to 5
set keep with next of paragraph format of semanticStyle to true
set font size of font object of semanticStyle to 15
set name of font object of semanticStyle to "黑体"
set east asian name of font object of semanticStyle to "黑体"
set ascii name of font object of semanticStyle to "黑体"
set other name of font object of semanticStyle to "黑体"
set complex script name of font object of semanticStyle to "黑体"
set line spacing rule of paragraph format of semanticStyle to line space single
set semanticStyle to requestedHeading4
set bold of font object of semanticStyle to true
set alignment of paragraph format of semanticStyle to align paragraph left
set space before of paragraph format of semanticStyle to 10
set space after of paragraph format of semanticStyle to 5
set keep with next of paragraph format of semanticStyle to true
set font size of font object of semanticStyle to 14
set name of font object of semanticStyle to "黑体"
set east asian name of font object of semanticStyle to "黑体"
set ascii name of font object of semanticStyle to "黑体"
set other name of font object of semanticStyle to "黑体"
set complex script name of font object of semanticStyle to "黑体"
set line spacing rule of paragraph format of semanticStyle to line space single
set semanticStyle to requestedHeading5
set bold of font object of semanticStyle to true
set alignment of paragraph format of semanticStyle to align paragraph left
set space before of paragraph format of semanticStyle to 8
set space after of paragraph format of semanticStyle to 5
set keep with next of paragraph format of semanticStyle to true
set font size of font object of semanticStyle to 14
set name of font object of semanticStyle to "黑体"
set east asian name of font object of semanticStyle to "黑体"
set ascii name of font object of semanticStyle to "黑体"
set other name of font object of semanticStyle to "黑体"
set complex script name of font object of semanticStyle to "黑体"
set line spacing rule of paragraph format of semanticStyle to line space single
set semanticStyle to requestedHeading6
set bold of font object of semanticStyle to true
set alignment of paragraph format of semanticStyle to align paragraph left
set space before of paragraph format of semanticStyle to 6
set space after of paragraph format of semanticStyle to 5
set keep with next of paragraph format of semanticStyle to true
set font size of font object of semanticStyle to 12
set name of font object of semanticStyle to "黑体"
set east asian name of font object of semanticStyle to "黑体"
set ascii name of font object of semanticStyle to "黑体"
set other name of font object of semanticStyle to "黑体"
set complex script name of font object of semanticStyle to "黑体"
set line spacing rule of paragraph format of semanticStyle to line space single
set nativeRows to {{"ok"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
