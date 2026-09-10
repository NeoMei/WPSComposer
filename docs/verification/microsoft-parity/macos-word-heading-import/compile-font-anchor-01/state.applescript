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
set firstRange to text object of paragraph 1 of boundDoc
set secondRange to text object of paragraph 2 of boundDoc
set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,start of content of firstRange,end of content of firstRange,start of content of secondRange,end of content of secondRange,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc}}
end tell
return my jsonRows(nativeRows)
