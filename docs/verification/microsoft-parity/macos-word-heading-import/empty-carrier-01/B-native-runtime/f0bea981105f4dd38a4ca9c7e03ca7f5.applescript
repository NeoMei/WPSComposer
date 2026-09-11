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
set boundDoc to document "document-c7c2012fc51544a0b28372d02e9d4423.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-7ejcbow5/document-c7c2012fc51544a0b28372d02e9d4423.docx" then error "WPSC_STALE_DOCUMENT"
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"SOURCE APPEARANCE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_FULL_PREIMAGE"
set sourceStyle to Word style "WPSC Heading Clone" of boundDoc
set cloneStyle to Word style "WPSC Heading Clone" of boundDoc
set sourceName to name local of sourceStyle as text
set cloneName to name local of cloneStyle as text
set nativeRows to {{"styles",sourceName,cloneName}}
set observedFont to font object of sourceStyle
set observedName to sourceName
set observedText to ""
set observedStart to -1
set observedEnd to -1
set end of nativeRows to {"font","source","style",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 32 end 66
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 32 or observedEnd is not 66 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"IMPORTED APPEARANCE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","paragraph",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 38 end 39
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 38 or observedEnd is not 39 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"E") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","latin",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 52 end 53
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 52 or observedEnd is not 53 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"中") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","cjk",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 55 end 56
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 55 or observedEnd is not 56 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"ا") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","arabic",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 63 end 65
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 63 or observedEnd is not 65 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"😀") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","emoji",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set observedFont to font object of cloneStyle
set observedName to cloneName
set observedText to ""
set observedStart to -1
set observedEnd to -1
set end of nativeRows to {"font","clone","style",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 73 end 99
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 73 or observedEnd is not 99 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","paragraph",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 75 end 76
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 75 or observedEnd is not 76 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"E") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","latin",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 85 end 86
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 85 or observedEnd is not 86 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"中") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","cjk",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 88 end 89
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 88 or observedEnd is not 89 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"ا") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","arabic",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 96 end 98
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 96 or observedEnd is not 98 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"😀") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","emoji",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
