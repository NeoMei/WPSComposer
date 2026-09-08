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
with timeout of 158 seconds
 tell application "Microsoft PowerPoint"
set identityCount to 0
repeat with di from 1 to (count of presentations)
 if name of presentation di is "bound-e02078bc209140fb86beaacbdf28a038.pptx" then set identityCount to identityCount + 1
end repeat
if identityCount is not 1 then error "Stale or ambiguous bound presentation"
set ownedDoc to presentation "bound-e02078bc209140fb86beaacbdf28a038.pptx"
if full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wpscomposer/session-_qei99bt/bound-e02078bc209140fb86beaacbdf28a038.pptx" then error "Bound presentation path changed"
set targetSlide to slide of view of document window 1 of ownedDoc
set layoutShape to make new line shape at end of targetSlide with properties {name:"wpscomposer-700545fbbd6a4f52b44bfa27c1147888",begin line X:40,begin line Y:412.0,end line X:490,end line Y:412.0}
set line weight of line format of layoutShape to 4
set fore color of line format of layoutShape to {170, 68, 34}
set layoutShape to make new shape at end of targetSlide with properties {name:"wpscomposer-28ae784558434ef9977d1eed1edc3549",left position:560,top:80,width:220,height:120,auto shape type:autoshape rectangle}
set fore color of fill format of layoutShape to {238, 204, 136}
set transparency of line format of layoutShape to 1
set content of text range of text frame of layoutShape to "Native box"
set font size of font of text range of text frame of layoutShape to 12
set layoutShape to make new text box at end of targetSlide with properties {name:"wpscomposer-f32cfa581b5a42f58b0118e531d2f9c6",left position:560,top:240,width:300,height:65}
set content of text range of text frame of layoutShape to "Native layout text"
set font size of font of text range of text frame of layoutShape to 22
set font color of font of text range of text frame of layoutShape to {18, 52, 86}
set bold of font of text range of text frame of layoutShape to true
set alignment of paragraph format of text range of text frame of layoutShape to paragraph align center
return "LAYOUT"
 end tell
end timeout
end sessionOperation
set operationResult to my sessionOperation()
return my encodeJSON({"WPSCOMPOSER_PPT_SESSION_OK",operationResult})
