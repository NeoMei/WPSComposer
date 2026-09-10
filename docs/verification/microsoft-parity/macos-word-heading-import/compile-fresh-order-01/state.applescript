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
set referenceRange to text object of paragraph 2 of boundDoc
set freshRange to text object of paragraph 4 of boundDoc
set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,start of content of referenceRange,end of content of referenceRange,start of content of freshRange,end of content of freshRange,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc}}
end tell
return my jsonRows(nativeRows)
