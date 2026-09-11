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
with timeout of 108 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-0cf183ed375548609d9539d2d9654522.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-0cf183ed375548609d9539d2d9654522.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-b531qa25/bound-0cf183ed375548609d9539d2d9654522.pptx" then error "Bound presentation path changed"
set targetSlide to slide 1 of ownedDoc
set originalSlideIds to {}
repeat with nativeIndex from 1 to count of slides of ownedDoc
 set end of originalSlideIds to slide ID of slide nativeIndex of ownedDoc
end repeat
copy object targetSlide
paste object ownedDoc
set copiedIndices to {}
repeat with nativeIndex from 1 to count of slides of ownedDoc
 if slide ID of slide nativeIndex of ownedDoc is not in originalSlideIds then set end of copiedIndices to nativeIndex as integer
end repeat
if count of copiedIndices is not 1 then error "Native slide paste count mismatch"
set targetSlide to slide (item 1 of copiedIndices) of ownedDoc
set destinationIndex to 1
if destinationIndex > count of slides of ownedDoc then set destinationIndex to count of slides of ownedDoc
set sourceIndex to slide index of targetSlide
if sourceIndex < destinationIndex then
 move targetSlide to after slide destinationIndex of ownedDoc
else if sourceIndex > destinationIndex then
 move targetSlide to before slide destinationIndex of ownedDoc
end if
return slide index of targetSlide
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
