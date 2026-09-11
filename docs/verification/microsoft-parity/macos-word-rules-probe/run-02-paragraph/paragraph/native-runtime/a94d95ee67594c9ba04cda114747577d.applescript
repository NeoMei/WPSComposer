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
with timeout of 60 seconds
tell application "/Applications/Microsoft Word.app"
set nativeRows to {}
set boundDoc to document "document-695a6e5581c2412c96f477f787409b74.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-h6yd3xd6/document-695a6e5581c2412c96f477f787409b74.docx" then error "WPSC_STALE_DOCUMENT"
set p to (end of content of text object of boundDoc) - 1
set insertionRange to create range boundDoc start p end p
set content of insertionRange to " " & return & "FOLLOWING"
set ruleRange to create range boundDoc start p end (p + 2)
set alignment of paragraph format of ruleRange to align paragraph center
set ownBorder to get border ruleRange which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {49344,49344,49344}
set followingRange to create range boundDoc start (p + 2) end (p + 11)
set followingBorder to get border followingRange which border border bottom
set line style of followingBorder to line style none
set ruleRange to text object of paragraph 2 of boundDoc
set followingRange to text object of paragraph 3 of boundDoc
set ownBorder to get border ruleRange which border border bottom
set followingBorder to get border followingRange which border border bottom
set nativeRows to {{"paragraph",content of ruleRange as text,(alignment of paragraph format of ruleRange is align paragraph center),(line style of ownBorder is line style single),(line width of ownBorder is line width75 point),(color of ownBorder is {49344,49344,49344}),content of followingRange as text,(line style of followingBorder is line style none)}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
