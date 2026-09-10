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
tell application "Microsoft Word"
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"SOURCE APPEARANCE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "" & return & "") as boolean) then error "WPSC_FONT_FULL_PREIMAGE"
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
set targetRange to create range boundDoc start 0 end 32
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 0 or observedEnd is not 32 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"SOURCE APPEARANCE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","paragraph",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 7 end 8
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 7 or observedEnd is not 8 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"A") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","latin",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 18 end 19
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 18 or observedEnd is not 19 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"中") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","cjk",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 21 end 22
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 21 or observedEnd is not 22 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"ا") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not sourceName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","source","arabic",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 29 end 31
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 29 or observedEnd is not 31 then error "WPSC_FONT_ANCHOR_BOUNDS"
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
set targetRange to create range boundDoc start 32 end 66
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 32 or observedEnd is not 66 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"IMPORTED APPEARANCE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","paragraph",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 41 end 42
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 41 or observedEnd is not 42 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"A") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","latin",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 52 end 53
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 52 or observedEnd is not 53 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"中") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","cjk",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 55 end 56
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 55 or observedEnd is not 56 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"ا") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","arabic",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
set targetRange to create range boundDoc start 63 end 65
set observedStart to start of content of targetRange
set observedEnd to end of content of targetRange
if observedStart is not 63 or observedEnd is not 65 then error "WPSC_FONT_ANCHOR_BOUNDS"
set observedText to content of targetRange as text
if not (((current application's NSString's stringWithString:observedText)'s isEqualToString:"😀") as boolean) then error "WPSC_FONT_ANCHOR_TEXT"
set observedName to name local of style of targetRange as text
if observedName is not cloneName then error "WPSC_FONT_ANCHOR_STYLE"
set observedFont to font object of targetRange
set end of nativeRows to {"font","clone","emoji",observedText,observedStart,observedEnd,observedName,ascii name of observedFont as text,complex script name of observedFont as text,name of observedFont as text,other name of observedFont as text,east asian name of observedFont as text}
end tell
return my jsonRows(nativeRows)
