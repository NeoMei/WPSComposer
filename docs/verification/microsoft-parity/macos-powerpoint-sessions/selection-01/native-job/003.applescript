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
with timeout of 97 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-dd3fba8e3bdf4eeabbc2c2ab4e619e9a.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-dd3fba8e3bdf4eeabbc2c2ab4e619e9a.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-p__56_3n/bound-dd3fba8e3bdf4eeabbc2c2ab4e619e9a.pptx" then error "Bound presentation path changed"
if full name of active presentation is not full name of ownedDoc then error "Selection belongs to a different presentation"
set sel to selection of document window 1 of ownedDoc
if selection type of sel is selection type text then
 set tx to text range of sel
 set fn to font of tx
 set pf to paragraph format of tx
 return my encodeJSON({"text",content of tx,font name of fn,font size of fn,bold of fn,italic of fn,underline of fn,font color of fn,alignment of pf as text,space before of pf,space after of pf,space within of pf})
end if
if selection type of sel is selection type none then return my encodeJSON({"none"})
if selection type of sel is selection type shapes then
 if has child shape range of sel then error "Child group shape selection is unsupported"
 set selectedIndices to {}
 set selectedShapes to shape range of sel
 repeat with selectedIndex from 1 to count of shapes of selectedShapes
  set end of selectedIndices to z order position of shape selectedIndex of selectedShapes
 end repeat
 return my encodeJSON({"shapes",slide index of slide of view of document window 1 of ownedDoc,selectedIndices})
end if
return my encodeJSON({"unverified",selection type of sel as text})
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
