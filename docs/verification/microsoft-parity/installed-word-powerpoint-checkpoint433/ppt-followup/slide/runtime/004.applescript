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
set pdfHFS to (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-5_hn6aav/export-a2b84cfc9777417cae424ec1fddb3491.pdf") as text
with timeout of 41 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-a916ecf6b56d483cb35bd84559ed7fcc.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-a916ecf6b56d483cb35bd84559ed7fcc.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-5_hn6aav/bound-a916ecf6b56d483cb35bd84559ed7fcc.pptx" then error "Bound presentation path changed"
save ownedDoc in pdfHFS as save as PDF
return "PDF"
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
