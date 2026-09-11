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
with timeout of 20 seconds
 tell application "Microsoft PowerPoint"
  set rows to {}
  repeat with di from 1 to count of presentations
   set allText to {}
   repeat with si from 1 to count of slides of presentation di
    repeat with sh from 1 to count of shapes of slide si of presentation di
     if has text frame of shape sh of slide si of presentation di then set end of allText to content of text range of text frame of shape sh of slide si of presentation di
    end repeat
   end repeat
   set end of rows to {name of presentation di,full name of presentation di,saved of presentation di,count of slides of presentation di,allText}
  end repeat
  return my encodeJSON(rows)
 end tell
end timeout