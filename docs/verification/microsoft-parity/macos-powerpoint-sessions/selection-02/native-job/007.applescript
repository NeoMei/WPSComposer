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
on sessionOperation()
with timeout of 82 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-a029ea3687cb477b8729b1b861e5a5ec.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-a029ea3687cb477b8729b1b861e5a5ec.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-5q4zj7q9/bound-a029ea3687cb477b8729b1b861e5a5ec.pptx" then error "Bound presentation path changed"
if full name of active presentation is not full name of ownedDoc then error "Selection changed before mutation"
set boundSelection to selection of document window 1 of ownedDoc
if selection type of boundSelection is not selection type shapes then error "Selection changed before mutation"
if count of shapes of shape range of boundSelection is not 1 then error "Selection changed before mutation"
if slide index of slide of view of document window 1 of ownedDoc is not 2 then error "Selection changed before mutation"
if z order position of shape 1 of shape range of boundSelection is not 2 then error "Selection changed before mutation"
set targetSlide to slide 2 of ownedDoc
set matchingShapes to {}
repeat with shapeIndex from 1 to (count of shapes of targetSlide)
 if name of shape shapeIndex of targetSlide is "parity-title" then set end of matchingShapes to shapeIndex
end repeat
if (count of matchingShapes) is not 1 then error "Stale or ambiguous shape name"
set targetShape to shape (item 1 of matchingShapes) of targetSlide
set targetText to text range of text frame of targetShape
set content of targetText to "Selected native patch"
set font size of font of targetText to 24
set left position of targetShape to 80
return "PATCHED"
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
