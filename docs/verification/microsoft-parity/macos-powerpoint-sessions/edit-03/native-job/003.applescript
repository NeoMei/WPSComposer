use framework "Foundation"
use scripting additions
on normalizedJSON(v)
 if v is missing value then return "__NATIVE_UNAVAILABLE__"
 if class of v is list then
  set resultArray to current application's NSMutableArray's array()
  repeat with itemRef in v
   set normalizedItem to my normalizedJSON(contents of itemRef)
   resultArray's addObject:normalizedItem
  end repeat
  return resultArray
 end if
 if class of v is reference then return my normalizedJSON(contents of v)
 if class of v is integer or class of v is real or class of v is boolean or class of v is text then return v
 return v as text
end normalizedJSON
on encodeJSON(itemsList)
 set normalizedItems to my normalizedJSON(itemsList)
 set jsonData to current application's NSJSONSerialization's dataWithJSONObject:normalizedItems options:0 |error|:(missing value)
 if jsonData is missing value then error "Native snapshot JSON encoding failed"
 return (current application's NSString's alloc()'s initWithData:jsonData encoding:4) as text
end encodeJSON
with timeout of 115 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-b7d3bcabd4a842278647ea6aaaef690f.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-b7d3bcabd4a842278647ea6aaaef690f.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-rbcwpomp/bound-b7d3bcabd4a842278647ea6aaaef690f.pptx" then error "Bound presentation path changed"
set targetSlide to slide 1 of ownedDoc
set matchingShapes to {}
repeat with shapeIndex from 1 to (count of shapes of targetSlide)
 if name of shape shapeIndex of targetSlide is "parity-title" then set end of matchingShapes to shapeIndex
end repeat
if (count of matchingShapes) is not 1 then error "Stale or ambiguous shape name"
set targetShape to shape (item 1 of matchingShapes) of targetSlide
set targetText to text range of text frame of targetShape
set content of targetText to "Session native edit"
set font name of font of targetText to "Arial"
set east asian name of font of targetText to "Arial"
set font size of font of targetText to 26
set bold of font of targetText to true
set italic of font of targetText to true
set underline of font of targetText to true
set font color of font of targetText to {18, 52, 86}
set alignment of paragraph format of targetText to paragraph align center
set space before of paragraph format of targetText to 3
set space after of paragraph format of targetText to 6
set space within of paragraph format of targetText to 1.2
set transparency of fill format of targetShape to 0
set transparency of line format of targetShape to 0
set margin left of text frame of targetShape to 8
set margin right of text frame of targetShape to 9
set margin top of text frame of targetShape to 10
set margin bottom of text frame of targetShape to 11
set word wrap of text frame of targetShape to true
set auto size of text frame of targetShape to auto size none
set vertical anchor of text frame of targetShape to anchor middle
set visible of fill format of targetShape to true
set line weight of line format of targetShape to 2
set fore color of fill format of targetShape to {241, 226, 211}
set fore color of line format of targetShape to {68, 85, 102}
set left position of targetShape to 45
set top of targetShape to 50
set width of targetShape to 550
set height of targetShape to 70
return "PATCHED"
 end tell
end timeout
