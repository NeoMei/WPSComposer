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
with timeout of 177 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-3ab0008656a84b16b88b3b720182e76a.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-3ab0008656a84b16b88b3b720182e76a.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-70uz4mj2/bound-3ab0008656a84b16b88b3b720182e76a.pptx" then error "Bound presentation path changed"
set inventoryRows to {}
repeat with di from 1 to count presentations
 set d to presentation di
 if full name of d is not full name of ownedDoc then
  set allText to ""
  repeat with si from 1 to count slides of d
   repeat with sh from 1 to count shapes of slide si of d
    if has text frame of shape sh of slide si of d then set allText to allText & (content of text range of text frame of shape sh of slide si of d as text) & return
   end repeat
  end repeat
  set textHash to do shell script ("/usr/bin/printf %s " & quoted form of allText & " | /usr/bin/shasum -a 256")
  set end of inventoryRows to {name of d,full name of d,saved of d,count slides of d,textHash}
 end if
end repeat
return my encodeJSON(inventoryRows)
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
