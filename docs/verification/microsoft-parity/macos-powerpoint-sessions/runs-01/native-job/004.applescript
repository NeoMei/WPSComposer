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
with timeout of 94 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-5d46fa8ae45b4fdebdc7bbb8e4b3817a.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-5d46fa8ae45b4fdebdc7bbb8e4b3817a.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-si0t0ie1/bound-5d46fa8ae45b4fdebdc7bbb8e4b3817a.pptx" then error "Bound presentation path changed"
set targetSlide to slide 1 of ownedDoc
set matchingShapes to {}
repeat with shapeIndex from 1 to (count of shapes of targetSlide)
 if name of shape shapeIndex of targetSlide is "parity-title" then set end of matchingShapes to shapeIndex
end repeat
if (count of matchingShapes) is not 1 then error "Stale or ambiguous shape name"
set targetShape to shape (item 1 of matchingShapes) of targetSlide
set targetText to text range of text frame of targetShape
set targetText to paragraph 1 of targetText
set targetText to text flow 2 of targetText
set content of targetText to "ession native run"
set strike type of font of targetText to single strike
set font color of font of targetText to {170, 17, 34}
return "PATCHED"
 end tell
end timeout
