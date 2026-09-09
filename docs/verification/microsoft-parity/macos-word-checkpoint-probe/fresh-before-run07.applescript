use framework "Foundation"
use scripting additions
on jsonRows(rows)
  set dataValue to current application's NSJSONSerialization's dataWithJSONObject:rows options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:dataValue encoding:4) as text
end jsonRows
on enumIndex(v, choices)
  repeat with i from 1 to count choices
    if v is item i of choices then return i - 1
  end repeat
  return -1
end enumIndex
set nativeRows to {}
if application "Microsoft Word" is running then
tell application "Microsoft Word"
repeat with di from 1 to count documents
set inventoryDoc to document di
set documentText to content of text object of inventoryDoc as text
set textHash to do shell script ("/usr/bin/printf %s " & quoted form of documentText & " | /usr/bin/shasum -a 256")
set end of nativeRows to {name of inventoryDoc as text,posix full name of inventoryDoc as text,saved of inventoryDoc,textHash}
end repeat
end tell
end if
return my jsonRows(nativeRows)