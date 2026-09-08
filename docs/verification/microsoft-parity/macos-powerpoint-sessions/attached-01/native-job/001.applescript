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
if application "Microsoft PowerPoint" is not running then error "PowerPoint is not running"
with timeout of 44 seconds
 tell application "Microsoft PowerPoint"
set activeDoc to active presentation
set activeName to name of activeDoc
set activePath to full name of activeDoc
set matched to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is activeName then set matched to matched + 1
end repeat
if matched is not 1 then error "Ambiguous active presentation"
return my encodeJSON({activeName,activePath,saved of activeDoc,read only of activeDoc})
 end tell
end timeout
