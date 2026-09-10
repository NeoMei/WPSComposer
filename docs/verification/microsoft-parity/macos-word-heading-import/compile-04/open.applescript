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

tell application "Microsoft Word"
set boundDoc to missing value
if (count documents) is not 0 then error "WPSC_FLATOPC_REQUIRES_EMPTY"
open file name "/tmp/flatopc-owned.xml" file converter open format xmldocument serialized read only true add to recent files false confirm conversions false
set matches to {}
repeat with nativeDocument in documents
if (((current application's NSString's stringWithString:(posix full name of nativeDocument as text))'s isEqualToString:"/tmp/flatopc-owned.xml") as boolean) then set end of matches to nativeDocument
end repeat
if (count matches) is not 1 or (count documents) is not 1 then error "WPSC_FLATOPC_IDENTITY_DELTA"
set boundDoc to item 1 of matches
set boundWindow to active window of boundDoc
if not (((current application's NSString's stringWithString:(name of boundDoc as text))'s isEqualToString:"flatopc-owned.xml") as boolean) then error "WPSC_FLATOPC_NAME"
set nativeRows to {{"binding",id of boundWindow,posix full name of boundDoc as text,name of boundDoc as text}}
end tell
