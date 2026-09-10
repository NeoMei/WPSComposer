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
set boundDoc to document "document-167d00e35a4246e098efd41109f7fc0a.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s2ag5496/document-167d00e35a4246e098efd41109f7fc0a.docx" then error "WPSC_STALE_DOCUMENT"
set sourceStyle to Word style (style heading1) of boundDoc
set detachedStyle to Word style "WPSC Heading Clone" of boundDoc
set sourceFont to font object of text object of paragraph 1 of boundDoc
set destinationFont to font object of text object of paragraph 4 of boundDoc
set sourceParagraph to paragraph format of sourceStyle
set destinationParagraph to paragraph format of detachedStyle
set nativeRows to {}
set leftValue to all caps of sourceFont
set rightValue to all caps of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:all caps",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to animation of sourceFont
set rightValue to animation of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:animation",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to ascii name of sourceFont
set rightValue to ascii name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:ascii name",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to bold of sourceFont
set rightValue to bold of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:bold",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to bold bi of sourceFont
set rightValue to bold bi of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:bold bi",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to color of sourceFont
set rightValue to color of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:color",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to color index of sourceFont
set rightValue to color index of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:color index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to color theme index of sourceFont
set rightValue to color theme index of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:color theme index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to complex script name of sourceFont
set rightValue to complex script name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:complex script name",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to disable character space grid of sourceFont
set rightValue to disable character space grid of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:disable character space grid",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to double strike through of sourceFont
set rightValue to double strike through of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:double strike through",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to east asian name of sourceFont
set rightValue to east asian name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:east asian name",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to emboss of sourceFont
set rightValue to emboss of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:emboss",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to emphasis mark of sourceFont
set rightValue to emphasis mark of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:emphasis mark",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to engrave of sourceFont
set rightValue to engrave of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:engrave",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to font position of sourceFont
set rightValue to font position of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:font position",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to font size of sourceFont
set rightValue to font size of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:font size",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to hidden of sourceFont
set rightValue to hidden of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:hidden",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to italic of sourceFont
set rightValue to italic of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:italic",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to italic bi of sourceFont
set rightValue to italic bi of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:italic bi",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to kerning of sourceFont
set rightValue to kerning of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:kerning",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to name of sourceFont
set rightValue to name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:name",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to other name of sourceFont
set rightValue to other name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:other name",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to outline of sourceFont
set rightValue to outline of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:outline",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to scaling of sourceFont
set rightValue to scaling of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:scaling",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to shadow of sourceFont
set rightValue to shadow of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:shadow",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to small caps of sourceFont
set rightValue to small caps of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:small caps",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to spacing of sourceFont
set rightValue to spacing of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:spacing",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to strike through of sourceFont
set rightValue to strike through of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:strike through",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to subscript of sourceFont
set rightValue to subscript of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:subscript",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to superscript of sourceFont
set rightValue to superscript of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:superscript",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to underline of sourceFont
set rightValue to underline of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:underline",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to underline color of sourceFont
set rightValue to underline color of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:underline color",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to underline color theme index of sourceFont
set rightValue to underline color theme index of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font:underline color theme index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to background pattern color of shading of sourceFont
set rightValue to background pattern color of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-shading:background pattern color",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to background pattern color index of shading of sourceFont
set rightValue to background pattern color index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-shading:background pattern color index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to background pattern color theme index of shading of sourceFont
set rightValue to background pattern color theme index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-shading:background pattern color theme index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to foreground pattern color of shading of sourceFont
set rightValue to foreground pattern color of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-shading:foreground pattern color",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to foreground pattern color index of shading of sourceFont
set rightValue to foreground pattern color index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-shading:foreground pattern color index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to foreground pattern color theme index of shading of sourceFont
set rightValue to foreground pattern color theme index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-shading:foreground pattern color theme index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to texture of shading of sourceFont
set rightValue to texture of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-shading:texture",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to always in front of border options of sourceFont
set rightValue to always in front of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:always in front",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to distance from of border options of sourceFont
set rightValue to distance from of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:distance from",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to distance from bottom of border options of sourceFont
set rightValue to distance from bottom of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:distance from bottom",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to distance from left of border options of sourceFont
set rightValue to distance from left of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:distance from left",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to distance from right of border options of sourceFont
set rightValue to distance from right of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:distance from right",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to distance from top of border options of sourceFont
set rightValue to distance from top of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:distance from top",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to enable borders of border options of sourceFont
set rightValue to enable borders of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:enable borders",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to enable first page in section of border options of sourceFont
set rightValue to enable first page in section of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:enable first page in section",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to enable other pages in section of border options of sourceFont
set rightValue to enable other pages in section of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:enable other pages in section",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to inside color of border options of sourceFont
set rightValue to inside color of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:inside color",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to inside color index of border options of sourceFont
set rightValue to inside color index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:inside color index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to inside color theme index of border options of sourceFont
set rightValue to inside color theme index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:inside color theme index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to inside line style of border options of sourceFont
set rightValue to inside line style of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:inside line style",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to inside line width of border options of sourceFont
set rightValue to inside line width of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:inside line width",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to join borders of border options of sourceFont
set rightValue to join borders of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:join borders",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to outside color of border options of sourceFont
set rightValue to outside color of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:outside color",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to outside color index of border options of sourceFont
set rightValue to outside color index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:outside color index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to outside color theme index of border options of sourceFont
set rightValue to outside color theme index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:outside color theme index",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to outside line style of border options of sourceFont
set rightValue to outside line style of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:outside line style",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to outside line width of border options of sourceFont
set rightValue to outside line width of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:outside line width",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to shadow of border options of sourceFont
set rightValue to shadow of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:shadow",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to surround footer of border options of sourceFont
set rightValue to surround footer of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:surround footer",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to surround header of border options of sourceFont
set rightValue to surround header of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"font-borders:surround header",{leftKind,leftValue},{rightKind,rightValue},sameValue}
set leftValue to name local of Word style (style normal) of boundDoc as text
set rightValue to name local of Word style (base style of detachedStyle) of boundDoc as text
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
if class of leftValue is text then
set leftKind to "text"
else if class of leftValue is boolean then
set leftKind to "boolean"
else if (class of leftValue is integer) or (class of leftValue is real) then
set leftKind to "number"
else if class of leftValue is list then
set leftKind to "number-list"
else
set leftKind to "native-enum"
set leftValue to leftValue as text
end if
if class of rightValue is text then
set rightKind to "text"
else if class of rightValue is boolean then
set rightKind to "boolean"
else if (class of rightValue is integer) or (class of rightValue is real) then
set rightKind to "number"
else if class of rightValue is list then
set rightKind to "number-list"
else
set rightKind to "native-enum"
set rightValue to rightValue as text
end if
set end of nativeRows to {"clone-base-normal",{leftKind,leftValue},{rightKind,rightValue},sameValue}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
