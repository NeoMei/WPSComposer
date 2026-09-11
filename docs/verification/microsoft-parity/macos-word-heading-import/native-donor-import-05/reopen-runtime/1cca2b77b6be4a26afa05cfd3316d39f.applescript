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
set boundDoc to document "document-e229b8af61d64aeabf3e6a30ff701c5d.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-p31smp36/document-e229b8af61d64aeabf3e6a30ff701c5d.docx" then error "WPSC_STALE_DOCUMENT"
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_FULL_PREIMAGE"
set sourceStyle to Word style (style heading1) of boundDoc
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
set targetRange to create range boundDoc start 0 end 26
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 0 or observedEnd is not 26 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","paragraph",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 2 end 3
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 2 or observedEnd is not 3 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"E") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","latin",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 12 end 13
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 12 or observedEnd is not 13 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"中") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","cjk",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 15 end 16
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 15 or observedEnd is not 16 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"ا") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","arabic",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 23 end 25
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 23 or observedEnd is not 25 then error "WPSC_FONT_ANCHOR_BOUNDS"
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
set targetRange to create range boundDoc start 67 end 93
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 67 or observedEnd is not 93 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","paragraph",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 69 end 70
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 69 or observedEnd is not 70 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"E") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","latin",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 79 end 80
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 79 or observedEnd is not 80 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"中") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","cjk",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 82 end 83
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 82 or observedEnd is not 83 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"ا") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","arabic",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 90 end 92
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 90 or observedEnd is not 92 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"😀") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","emoji",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
