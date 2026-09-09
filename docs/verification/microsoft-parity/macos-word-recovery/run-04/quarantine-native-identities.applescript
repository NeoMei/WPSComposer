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
set end of nativeRows to {"counts",count documents,count windows}
repeat with di from 1 to count documents
try
set d to document di
set end of nativeRows to {"doc",di as integer,name of d as text,posix full name of d as text,saved of d}
on error errorText number errorNumber
set end of nativeRows to {"doc-error",di as integer,errorNumber,errorText}
end try
end repeat
repeat with wi from 1 to count windows
try
set w to window wi
set end of nativeRows to {"window",wi as integer,name of w as text,name of document of w as text}
on error errorText number errorNumber
set end of nativeRows to {"window-error",wi as integer,errorNumber,errorText}
end try
end repeat
end tell
end if
return my jsonRows(nativeRows)