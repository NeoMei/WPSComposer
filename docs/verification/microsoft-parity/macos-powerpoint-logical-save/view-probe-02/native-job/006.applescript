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
with timeout of 170 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-eab8809fe8a94b1b8ff1acfbcd49d7e3.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-eab8809fe8a94b1b8ff1acfbcd49d7e3.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-g4w1qdcc/bound-eab8809fe8a94b1b8ff1acfbcd49d7e3.pptx" then error "Bound presentation path changed"
set results to {}
try
set index of window (name of ownedDoc) to 1
set end of results to {"window-index",true,full name of active presentation is full name of ownedDoc}
on error errorText number errorNumber
set end of results to {"window-index",false,errorNumber}
end try
try
set visible of window (name of ownedDoc) to true
set end of results to {"window-visible",true,full name of active presentation is full name of ownedDoc}
on error errorText number errorNumber
set end of results to {"window-visible",false,errorNumber}
end try
try
activate ownedDoc
set end of results to {"activate-presentation",true,full name of active presentation is full name of ownedDoc}
on error errorText number errorNumber
set end of results to {"activate-presentation",false,errorNumber}
end try
try
set active of document window 1 of ownedDoc to true
set end of results to {"set-window-active",true,full name of active presentation is full name of ownedDoc}
on error errorText number errorNumber
set end of results to {"set-window-active",false,errorNumber}
end try
try
open (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-g4w1qdcc/bound-eab8809fe8a94b1b8ff1acfbcd49d7e3.pptx" as alias)
set end of results to {"open-exact-private",true,full name of active presentation is full name of ownedDoc}
on error errorText number errorNumber
set end of results to {"open-exact-private",false,errorNumber}
end try
try
go to slide (view of document window 1 of ownedDoc) number 2
set end of results to {"goto-owned",true,slide index of slide of view of document window 1 of ownedDoc}
on error errorText number errorNumber
set end of results to {"goto-owned",false,errorNumber}
end try
return my encodeJSON(results)
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
