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
set boundDoc to document "document-7502140126cb4c4798e75ae2ad5c1e7c.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-rarvrbrt/document-7502140126cb4c4798e75ae2ad5c1e7c.docx" then error "WPSC_STALE_DOCUMENT"
set ruleRange to text object of paragraph 1 of boundDoc
set followingRange to text object of paragraph 2 of boundDoc
set ownBorder to get border ruleRange which border border bottom
set followingBorder to get border followingRange which border border bottom
set fmt to paragraph format of followingRange
set followingStyleName to name local of style of followingRange as text
set expectedBodyStyleName to name local of Word style (style body text) of boundDoc as text
set followingIsBodyText to ((current application's NSString's stringWithString:followingStyleName)'s isEqualToString:expectedBodyStyleName) as boolean
set nativeRows to {{"public-paragraph-rule",content of ruleRange as text,content of followingRange as text,(alignment of paragraph format of ruleRange is align paragraph center),(line style of ownBorder is line style single),(line width of ownBorder is line width75 point),(color of ownBorder is {49344,49344,49344}),(line style of followingBorder is line style none),followingIsBodyText,(alignment of fmt is align paragraph center),space before of fmt,space after of fmt,first line indent of fmt,line spacing of fmt}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
