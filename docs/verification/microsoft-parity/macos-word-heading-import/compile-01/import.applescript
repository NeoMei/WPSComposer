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
set importRange to text object of paragraph 2 of boundDoc
set importStart to start of content of importRange
set importEnd to end of content of importRange
if content of importRange is not "REPLACE" & return then error "WPSC_IMPORT_PREIMAGE"
insert file at importRange file name "/tmp/owned-heading-input.xml" confirm conversions false link false
set nativeRows to {{"import",importStart,importEnd,content of text object of boundDoc as text}}
end tell
