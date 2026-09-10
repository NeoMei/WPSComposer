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
set sourceParagraph to paragraph format of sourceStyle
set cloneParagraph to paragraph format of cloneStyle
set sourceTabs to (get every tab stop of sourceParagraph)
set cloneTabs to (get every tab stop of cloneParagraph)
if class of sourceTabs is not list or class of cloneTabs is not list then error "WPSC_TAB_LIST_TYPE"
set sourceCount to count sourceTabs
set cloneCount to count cloneTabs
set nativeRows to {{"tabs","style",sourceName,cloneName,sourceCount,cloneCount}}
repeat with tabIndex from 1 to sourceCount
set ownTab to item tabIndex of sourceTabs
set end of nativeRows to {"source",tabIndex,tab stop position of ownTab,alignment of ownTab as text,tab leader of ownTab as text,custom tab of ownTab}
end repeat
repeat with tabIndex from 1 to cloneCount
set ownTab to item tabIndex of cloneTabs
set end of nativeRows to {"clone",tabIndex,tab stop position of ownTab,alignment of ownTab as text,tab leader of ownTab as text,custom tab of ownTab}
end repeat
end tell
return my jsonRows(nativeRows)
