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
set sourceStyle to Word style (style heading1) of boundDoc
set detachedStyle to Word style "WPSC Heading Clone" of boundDoc
set sourceFont to font object of sourceStyle
set destinationFont to font object of detachedStyle
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
set end of nativeRows to {"font:all caps",sameValue}
set leftValue to animation of sourceFont
set rightValue to animation of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:animation",sameValue}
set leftValue to ascii name of sourceFont
set rightValue to ascii name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:ascii name",sameValue}
set leftValue to bold of sourceFont
set rightValue to bold of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:bold",sameValue}
set leftValue to bold bi of sourceFont
set rightValue to bold bi of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:bold bi",sameValue}
set leftValue to color of sourceFont
set rightValue to color of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:color",sameValue}
set leftValue to color index of sourceFont
set rightValue to color index of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:color index",sameValue}
set leftValue to color theme index of sourceFont
set rightValue to color theme index of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:color theme index",sameValue}
set leftValue to complex script name of sourceFont
set rightValue to complex script name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:complex script name",sameValue}
set leftValue to disable character space grid of sourceFont
set rightValue to disable character space grid of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:disable character space grid",sameValue}
set leftValue to double strike through of sourceFont
set rightValue to double strike through of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:double strike through",sameValue}
set leftValue to east asian name of sourceFont
set rightValue to east asian name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:east asian name",sameValue}
set leftValue to emboss of sourceFont
set rightValue to emboss of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:emboss",sameValue}
set leftValue to emphasis mark of sourceFont
set rightValue to emphasis mark of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:emphasis mark",sameValue}
set leftValue to engrave of sourceFont
set rightValue to engrave of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:engrave",sameValue}
set leftValue to font position of sourceFont
set rightValue to font position of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:font position",sameValue}
set leftValue to font size of sourceFont
set rightValue to font size of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:font size",sameValue}
set leftValue to hidden of sourceFont
set rightValue to hidden of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:hidden",sameValue}
set leftValue to italic of sourceFont
set rightValue to italic of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:italic",sameValue}
set leftValue to italic bi of sourceFont
set rightValue to italic bi of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:italic bi",sameValue}
set leftValue to kerning of sourceFont
set rightValue to kerning of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:kerning",sameValue}
set leftValue to name of sourceFont
set rightValue to name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:name",sameValue}
set leftValue to other name of sourceFont
set rightValue to other name of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:other name",sameValue}
set leftValue to outline of sourceFont
set rightValue to outline of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:outline",sameValue}
set leftValue to scaling of sourceFont
set rightValue to scaling of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:scaling",sameValue}
set leftValue to shadow of sourceFont
set rightValue to shadow of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:shadow",sameValue}
set leftValue to small caps of sourceFont
set rightValue to small caps of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:small caps",sameValue}
set leftValue to spacing of sourceFont
set rightValue to spacing of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:spacing",sameValue}
set leftValue to strike through of sourceFont
set rightValue to strike through of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:strike through",sameValue}
set leftValue to subscript of sourceFont
set rightValue to subscript of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:subscript",sameValue}
set leftValue to superscript of sourceFont
set rightValue to superscript of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:superscript",sameValue}
set leftValue to underline of sourceFont
set rightValue to underline of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:underline",sameValue}
set leftValue to underline color of sourceFont
set rightValue to underline color of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:underline color",sameValue}
set leftValue to underline color theme index of sourceFont
set rightValue to underline color theme index of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font:underline color theme index",sameValue}
set leftValue to background pattern color of shading of sourceFont
set rightValue to background pattern color of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-shading:background pattern color",sameValue}
set leftValue to background pattern color index of shading of sourceFont
set rightValue to background pattern color index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-shading:background pattern color index",sameValue}
set leftValue to background pattern color theme index of shading of sourceFont
set rightValue to background pattern color theme index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-shading:background pattern color theme index",sameValue}
set leftValue to foreground pattern color of shading of sourceFont
set rightValue to foreground pattern color of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-shading:foreground pattern color",sameValue}
set leftValue to foreground pattern color index of shading of sourceFont
set rightValue to foreground pattern color index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-shading:foreground pattern color index",sameValue}
set leftValue to foreground pattern color theme index of shading of sourceFont
set rightValue to foreground pattern color theme index of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-shading:foreground pattern color theme index",sameValue}
set leftValue to texture of shading of sourceFont
set rightValue to texture of shading of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-shading:texture",sameValue}
set leftValue to always in front of border options of sourceFont
set rightValue to always in front of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:always in front",sameValue}
set leftValue to distance from of border options of sourceFont
set rightValue to distance from of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:distance from",sameValue}
set leftValue to distance from bottom of border options of sourceFont
set rightValue to distance from bottom of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:distance from bottom",sameValue}
set leftValue to distance from left of border options of sourceFont
set rightValue to distance from left of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:distance from left",sameValue}
set leftValue to distance from right of border options of sourceFont
set rightValue to distance from right of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:distance from right",sameValue}
set leftValue to distance from top of border options of sourceFont
set rightValue to distance from top of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:distance from top",sameValue}
set leftValue to enable borders of border options of sourceFont
set rightValue to enable borders of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:enable borders",sameValue}
set leftValue to enable first page in section of border options of sourceFont
set rightValue to enable first page in section of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:enable first page in section",sameValue}
set leftValue to enable other pages in section of border options of sourceFont
set rightValue to enable other pages in section of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:enable other pages in section",sameValue}
set leftValue to inside color of border options of sourceFont
set rightValue to inside color of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:inside color",sameValue}
set leftValue to inside color index of border options of sourceFont
set rightValue to inside color index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:inside color index",sameValue}
set leftValue to inside color theme index of border options of sourceFont
set rightValue to inside color theme index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:inside color theme index",sameValue}
set leftValue to inside line style of border options of sourceFont
set rightValue to inside line style of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:inside line style",sameValue}
set leftValue to inside line width of border options of sourceFont
set rightValue to inside line width of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:inside line width",sameValue}
set leftValue to join borders of border options of sourceFont
set rightValue to join borders of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:join borders",sameValue}
set leftValue to outside color of border options of sourceFont
set rightValue to outside color of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:outside color",sameValue}
set leftValue to outside color index of border options of sourceFont
set rightValue to outside color index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:outside color index",sameValue}
set leftValue to outside color theme index of border options of sourceFont
set rightValue to outside color theme index of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:outside color theme index",sameValue}
set leftValue to outside line style of border options of sourceFont
set rightValue to outside line style of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:outside line style",sameValue}
set leftValue to outside line width of border options of sourceFont
set rightValue to outside line width of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:outside line width",sameValue}
set leftValue to shadow of border options of sourceFont
set rightValue to shadow of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:shadow",sameValue}
set leftValue to surround footer of border options of sourceFont
set rightValue to surround footer of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:surround footer",sameValue}
set leftValue to surround header of border options of sourceFont
set rightValue to surround header of border options of destinationFont
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"font-borders:surround header",sameValue}
set leftValue to add space between east asian and alpha of sourceParagraph
set rightValue to add space between east asian and alpha of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:add space between east asian and alpha",sameValue}
set leftValue to add space between east asian and digit of sourceParagraph
set rightValue to add space between east asian and digit of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:add space between east asian and digit",sameValue}
set leftValue to alignment of sourceParagraph
set rightValue to alignment of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:alignment",sameValue}
set leftValue to auto adjust right indent of sourceParagraph
set rightValue to auto adjust right indent of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:auto adjust right indent",sameValue}
set leftValue to base line alignment of sourceParagraph
set rightValue to base line alignment of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:base line alignment",sameValue}
set leftValue to character unit first line indent of sourceParagraph
set rightValue to character unit first line indent of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:character unit first line indent",sameValue}
set leftValue to character unit left indent of sourceParagraph
set rightValue to character unit left indent of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:character unit left indent",sameValue}
set leftValue to character unit right indent of sourceParagraph
set rightValue to character unit right indent of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:character unit right indent",sameValue}
set leftValue to disable line height grid of sourceParagraph
set rightValue to disable line height grid of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:disable line height grid",sameValue}
set leftValue to east asian line break control of sourceParagraph
set rightValue to east asian line break control of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:east asian line break control",sameValue}
set leftValue to first line indent of sourceParagraph
set rightValue to first line indent of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:first line indent",sameValue}
set leftValue to half width punctuation on top of line of sourceParagraph
set rightValue to half width punctuation on top of line of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:half width punctuation on top of line",sameValue}
set leftValue to hanging punctuation of sourceParagraph
set rightValue to hanging punctuation of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:hanging punctuation",sameValue}
set leftValue to hyphenation of sourceParagraph
set rightValue to hyphenation of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:hyphenation",sameValue}
set leftValue to keep together of sourceParagraph
set rightValue to keep together of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:keep together",sameValue}
set leftValue to keep with next of sourceParagraph
set rightValue to keep with next of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:keep with next",sameValue}
set leftValue to line spacing of sourceParagraph
set rightValue to line spacing of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:line spacing",sameValue}
set leftValue to line spacing rule of sourceParagraph
set rightValue to line spacing rule of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:line spacing rule",sameValue}
set leftValue to line unit after of sourceParagraph
set rightValue to line unit after of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:line unit after",sameValue}
set leftValue to line unit before of sourceParagraph
set rightValue to line unit before of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:line unit before",sameValue}
set leftValue to no line number of sourceParagraph
set rightValue to no line number of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:no line number",sameValue}
set leftValue to outline level of sourceParagraph
set rightValue to outline level of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:outline level",sameValue}
set leftValue to page break before of sourceParagraph
set rightValue to page break before of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:page break before",sameValue}
set leftValue to paragraph format left indent of sourceParagraph
set rightValue to paragraph format left indent of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:paragraph format left indent",sameValue}
set leftValue to paragraph format right indent of sourceParagraph
set rightValue to paragraph format right indent of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:paragraph format right indent",sameValue}
set leftValue to space after of sourceParagraph
set rightValue to space after of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:space after",sameValue}
set leftValue to space after auto of sourceParagraph
set rightValue to space after auto of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:space after auto",sameValue}
set leftValue to space before of sourceParagraph
set rightValue to space before of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:space before",sameValue}
set leftValue to space before auto of sourceParagraph
set rightValue to space before auto of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:space before auto",sameValue}
set leftValue to widow control of sourceParagraph
set rightValue to widow control of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:widow control",sameValue}
set leftValue to word wrap of sourceParagraph
set rightValue to word wrap of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph:word wrap",sameValue}
set leftValue to background pattern color of shading of sourceParagraph
set rightValue to background pattern color of shading of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-shading:background pattern color",sameValue}
set leftValue to background pattern color index of shading of sourceParagraph
set rightValue to background pattern color index of shading of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-shading:background pattern color index",sameValue}
set leftValue to background pattern color theme index of shading of sourceParagraph
set rightValue to background pattern color theme index of shading of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-shading:background pattern color theme index",sameValue}
set leftValue to foreground pattern color of shading of sourceParagraph
set rightValue to foreground pattern color of shading of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-shading:foreground pattern color",sameValue}
set leftValue to foreground pattern color index of shading of sourceParagraph
set rightValue to foreground pattern color index of shading of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-shading:foreground pattern color index",sameValue}
set leftValue to foreground pattern color theme index of shading of sourceParagraph
set rightValue to foreground pattern color theme index of shading of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-shading:foreground pattern color theme index",sameValue}
set leftValue to texture of shading of sourceParagraph
set rightValue to texture of shading of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-shading:texture",sameValue}
set leftValue to always in front of border options of sourceParagraph
set rightValue to always in front of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:always in front",sameValue}
set leftValue to distance from of border options of sourceParagraph
set rightValue to distance from of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:distance from",sameValue}
set leftValue to distance from bottom of border options of sourceParagraph
set rightValue to distance from bottom of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:distance from bottom",sameValue}
set leftValue to distance from left of border options of sourceParagraph
set rightValue to distance from left of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:distance from left",sameValue}
set leftValue to distance from right of border options of sourceParagraph
set rightValue to distance from right of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:distance from right",sameValue}
set leftValue to distance from top of border options of sourceParagraph
set rightValue to distance from top of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:distance from top",sameValue}
set leftValue to enable borders of border options of sourceParagraph
set rightValue to enable borders of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:enable borders",sameValue}
set leftValue to enable first page in section of border options of sourceParagraph
set rightValue to enable first page in section of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:enable first page in section",sameValue}
set leftValue to enable other pages in section of border options of sourceParagraph
set rightValue to enable other pages in section of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:enable other pages in section",sameValue}
set leftValue to inside color of border options of sourceParagraph
set rightValue to inside color of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:inside color",sameValue}
set leftValue to inside color index of border options of sourceParagraph
set rightValue to inside color index of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:inside color index",sameValue}
set leftValue to inside color theme index of border options of sourceParagraph
set rightValue to inside color theme index of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:inside color theme index",sameValue}
set leftValue to inside line style of border options of sourceParagraph
set rightValue to inside line style of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:inside line style",sameValue}
set leftValue to inside line width of border options of sourceParagraph
set rightValue to inside line width of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:inside line width",sameValue}
set leftValue to join borders of border options of sourceParagraph
set rightValue to join borders of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:join borders",sameValue}
set leftValue to outside color of border options of sourceParagraph
set rightValue to outside color of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:outside color",sameValue}
set leftValue to outside color index of border options of sourceParagraph
set rightValue to outside color index of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:outside color index",sameValue}
set leftValue to outside color theme index of border options of sourceParagraph
set rightValue to outside color theme index of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:outside color theme index",sameValue}
set leftValue to outside line style of border options of sourceParagraph
set rightValue to outside line style of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:outside line style",sameValue}
set leftValue to outside line width of border options of sourceParagraph
set rightValue to outside line width of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:outside line width",sameValue}
set leftValue to shadow of border options of sourceParagraph
set rightValue to shadow of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:shadow",sameValue}
set leftValue to surround footer of border options of sourceParagraph
set rightValue to surround footer of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:surround footer",sameValue}
set leftValue to surround header of border options of sourceParagraph
set rightValue to surround header of border options of destinationParagraph
if class of leftValue is text then
set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean
else
set sameValue to (leftValue is equal to rightValue)
end if
set end of nativeRows to {"paragraph-borders:surround header",sameValue}
set expectedBase to name local of Word style (style normal) of boundDoc as text
set observedBase to name local of Word style (base style of detachedStyle) of boundDoc as text
set baseEqual to ((current application's NSString's stringWithString:observedBase)'s isEqualToString:expectedBase) as boolean
set end of nativeRows to {"clone-base-normal",baseEqual}
end tell
return my jsonRows(nativeRows)
