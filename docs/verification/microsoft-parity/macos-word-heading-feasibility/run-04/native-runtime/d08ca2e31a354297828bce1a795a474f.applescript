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
set boundDoc to document "document-905fcb5e04b34279b257a112e505274f.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-6u_8q6g9/document-905fcb5e04b34279b257a112e505274f.docx" then error "WPSC_STALE_DOCUMENT"
set sourceStyle to Word style (style heading1) of boundDoc
set detachedStyle to Word style "WPSC Heading Clone" of boundDoc
set sourceFont to font object of sourceStyle
set destinationFont to font object of detachedStyle
set sourceParagraph to paragraph format of sourceStyle
set destinationParagraph to paragraph format of detachedStyle
set base style of detachedStyle to style normal
set scalarSourceValue to all caps of sourceFont
set scalarDestinationValue to all caps of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set all caps of destinationFont to scalarSourceValue
end if
set scalarSourceValue to animation of sourceFont
set scalarDestinationValue to animation of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set animation of destinationFont to scalarSourceValue
end if
set scalarSourceValue to ascii name of sourceFont
set scalarDestinationValue to ascii name of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set ascii name of destinationFont to scalarSourceValue
end if
set scalarSourceValue to bold of sourceFont
set scalarDestinationValue to bold of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set bold of destinationFont to scalarSourceValue
end if
set scalarSourceValue to bold bi of sourceFont
set scalarDestinationValue to bold bi of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set bold bi of destinationFont to scalarSourceValue
end if
set scalarSourceValue to color of sourceFont
set scalarDestinationValue to color of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set color of destinationFont to scalarSourceValue
end if
set scalarSourceValue to color index of sourceFont
set scalarDestinationValue to color index of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set color index of destinationFont to scalarSourceValue
end if
set scalarSourceValue to color theme index of sourceFont
set scalarDestinationValue to color theme index of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set color theme index of destinationFont to scalarSourceValue
end if
set scalarSourceValue to complex script name of sourceFont
set scalarDestinationValue to complex script name of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set complex script name of destinationFont to scalarSourceValue
end if
set scalarSourceValue to disable character space grid of sourceFont
set scalarDestinationValue to disable character space grid of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set disable character space grid of destinationFont to scalarSourceValue
end if
set scalarSourceValue to double strike through of sourceFont
set scalarDestinationValue to double strike through of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set double strike through of destinationFont to scalarSourceValue
end if
set scalarSourceValue to east asian name of sourceFont
set scalarDestinationValue to east asian name of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set east asian name of destinationFont to scalarSourceValue
end if
set scalarSourceValue to emboss of sourceFont
set scalarDestinationValue to emboss of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set emboss of destinationFont to scalarSourceValue
end if
set scalarSourceValue to emphasis mark of sourceFont
set scalarDestinationValue to emphasis mark of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set emphasis mark of destinationFont to scalarSourceValue
end if
set scalarSourceValue to engrave of sourceFont
set scalarDestinationValue to engrave of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set engrave of destinationFont to scalarSourceValue
end if
set scalarSourceValue to font position of sourceFont
set scalarDestinationValue to font position of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set font position of destinationFont to scalarSourceValue
end if
set scalarSourceValue to font size of sourceFont
set scalarDestinationValue to font size of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set font size of destinationFont to scalarSourceValue
end if
set scalarSourceValue to hidden of sourceFont
set scalarDestinationValue to hidden of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set hidden of destinationFont to scalarSourceValue
end if
set scalarSourceValue to italic of sourceFont
set scalarDestinationValue to italic of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set italic of destinationFont to scalarSourceValue
end if
set scalarSourceValue to italic bi of sourceFont
set scalarDestinationValue to italic bi of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set italic bi of destinationFont to scalarSourceValue
end if
set scalarSourceValue to kerning of sourceFont
set scalarDestinationValue to kerning of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set kerning of destinationFont to scalarSourceValue
end if
set scalarSourceValue to name of sourceFont
set scalarDestinationValue to name of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set name of destinationFont to scalarSourceValue
end if
set scalarSourceValue to other name of sourceFont
set scalarDestinationValue to other name of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set other name of destinationFont to scalarSourceValue
end if
set scalarSourceValue to outline of sourceFont
set scalarDestinationValue to outline of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set outline of destinationFont to scalarSourceValue
end if
set scalarSourceValue to scaling of sourceFont
set scalarDestinationValue to scaling of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set scaling of destinationFont to scalarSourceValue
end if
set scalarSourceValue to shadow of sourceFont
set scalarDestinationValue to shadow of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set shadow of destinationFont to scalarSourceValue
end if
set scalarSourceValue to small caps of sourceFont
set scalarDestinationValue to small caps of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set small caps of destinationFont to scalarSourceValue
end if
set scalarSourceValue to spacing of sourceFont
set scalarDestinationValue to spacing of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set spacing of destinationFont to scalarSourceValue
end if
set scalarSourceValue to strike through of sourceFont
set scalarDestinationValue to strike through of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set strike through of destinationFont to scalarSourceValue
end if
set scalarSourceValue to subscript of sourceFont
set scalarDestinationValue to subscript of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set subscript of destinationFont to scalarSourceValue
end if
set scalarSourceValue to superscript of sourceFont
set scalarDestinationValue to superscript of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set superscript of destinationFont to scalarSourceValue
end if
set scalarSourceValue to underline of sourceFont
set scalarDestinationValue to underline of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set underline of destinationFont to scalarSourceValue
end if
set scalarSourceValue to underline color of sourceFont
set scalarDestinationValue to underline color of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set underline color of destinationFont to scalarSourceValue
end if
set scalarSourceValue to underline color theme index of sourceFont
set scalarDestinationValue to underline color theme index of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set underline color theme index of destinationFont to scalarSourceValue
end if
set scalarSourceValue to background pattern color of shading of sourceFont
set scalarDestinationValue to background pattern color of shading of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set background pattern color of shading of destinationFont to scalarSourceValue
end if
set scalarSourceValue to background pattern color index of shading of sourceFont
set scalarDestinationValue to background pattern color index of shading of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set background pattern color index of shading of destinationFont to scalarSourceValue
end if
set scalarSourceValue to background pattern color theme index of shading of sourceFont
set scalarDestinationValue to background pattern color theme index of shading of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set background pattern color theme index of shading of destinationFont to scalarSourceValue
end if
set scalarSourceValue to foreground pattern color of shading of sourceFont
set scalarDestinationValue to foreground pattern color of shading of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set foreground pattern color of shading of destinationFont to scalarSourceValue
end if
set scalarSourceValue to foreground pattern color index of shading of sourceFont
set scalarDestinationValue to foreground pattern color index of shading of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set foreground pattern color index of shading of destinationFont to scalarSourceValue
end if
set scalarSourceValue to foreground pattern color theme index of shading of sourceFont
set scalarDestinationValue to foreground pattern color theme index of shading of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set foreground pattern color theme index of shading of destinationFont to scalarSourceValue
end if
set scalarSourceValue to texture of shading of sourceFont
set scalarDestinationValue to texture of shading of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set texture of shading of destinationFont to scalarSourceValue
end if
set scalarSourceValue to always in front of border options of sourceFont
set scalarDestinationValue to always in front of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set always in front of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to distance from of border options of sourceFont
set scalarDestinationValue to distance from of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set distance from of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to distance from bottom of border options of sourceFont
set scalarDestinationValue to distance from bottom of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set distance from bottom of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to distance from left of border options of sourceFont
set scalarDestinationValue to distance from left of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set distance from left of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to distance from right of border options of sourceFont
set scalarDestinationValue to distance from right of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set distance from right of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to distance from top of border options of sourceFont
set scalarDestinationValue to distance from top of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set distance from top of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to enable borders of border options of sourceFont
set scalarDestinationValue to enable borders of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set enable borders of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to enable first page in section of border options of sourceFont
set scalarDestinationValue to enable first page in section of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set enable first page in section of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to enable other pages in section of border options of sourceFont
set scalarDestinationValue to enable other pages in section of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set enable other pages in section of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to inside color of border options of sourceFont
set scalarDestinationValue to inside color of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set inside color of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to inside color index of border options of sourceFont
set scalarDestinationValue to inside color index of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set inside color index of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to inside color theme index of border options of sourceFont
set scalarDestinationValue to inside color theme index of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set inside color theme index of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to inside line style of border options of sourceFont
set scalarDestinationValue to inside line style of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set inside line style of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to inside line width of border options of sourceFont
set scalarDestinationValue to inside line width of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set inside line width of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to join borders of border options of sourceFont
set scalarDestinationValue to join borders of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set join borders of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to outside color of border options of sourceFont
set scalarDestinationValue to outside color of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set outside color of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to outside color index of border options of sourceFont
set scalarDestinationValue to outside color index of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set outside color index of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to outside color theme index of border options of sourceFont
set scalarDestinationValue to outside color theme index of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set outside color theme index of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to outside line style of border options of sourceFont
set scalarDestinationValue to outside line style of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set outside line style of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to outside line width of border options of sourceFont
set scalarDestinationValue to outside line width of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set outside line width of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to shadow of border options of sourceFont
set scalarDestinationValue to shadow of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set shadow of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to surround footer of border options of sourceFont
set scalarDestinationValue to surround footer of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set surround footer of border options of destinationFont to scalarSourceValue
end if
set scalarSourceValue to surround header of border options of sourceFont
set scalarDestinationValue to surround header of border options of destinationFont
if class of scalarSourceValue is text then
set scalarEqual to ((current application's NSString's stringWithString:scalarSourceValue)'s isEqualToString:scalarDestinationValue) as boolean
else
set scalarEqual to (scalarSourceValue is equal to scalarDestinationValue)
end if
if not scalarEqual then
set surround header of border options of destinationFont to scalarSourceValue
end if
set nativeRows to {{"stage",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
