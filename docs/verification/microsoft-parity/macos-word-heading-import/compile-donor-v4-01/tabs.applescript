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
set sourceStyle to Word style (style heading1) of boundDoc
set cloneStyle to Word style "WPSC Heading Clone" of boundDoc
set sourceName to name local of sourceStyle as text
set cloneName to name local of cloneStyle as text
set sourceRange to text object of paragraph 1 of boundDoc
set cloneRange to text object of paragraph 4 of boundDoc
if name local of style of sourceRange is not sourceName or name local of style of cloneRange is not cloneName then error "WPSC_TAB_PARAGRAPH_STYLE"
set sourceParagraph to paragraph 1 of sourceRange
set cloneParagraph to paragraph 1 of cloneRange
set sourceCount to count tab stops of sourceParagraph
set cloneCount to count tab stops of cloneParagraph
set nativeRows to {{"tabs","paragraph",sourceName,cloneName,sourceCount,cloneCount}}
repeat with tabIndex from 1 to sourceCount
set ownTab to tab stop tabIndex of sourceParagraph
set end of nativeRows to {"source",tabIndex,tab stop position of ownTab,alignment of ownTab as text,tab leader of ownTab as text,custom tab of ownTab}
end repeat
repeat with tabIndex from 1 to cloneCount
set ownTab to tab stop tabIndex of cloneParagraph
set end of nativeRows to {"clone",tabIndex,tab stop position of ownTab,alignment of ownTab as text,tab leader of ownTab as text,custom tab of ownTab}
end repeat
end tell
return my jsonRows(nativeRows)
