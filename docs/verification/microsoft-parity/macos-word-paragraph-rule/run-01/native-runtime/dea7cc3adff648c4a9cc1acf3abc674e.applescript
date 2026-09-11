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
set boundDoc to document "document-e5491f34ff764c00b74f3460d1308aac.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-bxlqsg9d/document-e5491f34ff764c00b74f3460d1308aac.docx" then error "WPSC_STALE_DOCUMENT"
set insertionPoint to (end of content of text object of boundDoc) - 1
set ruleOriginalPoint to insertionPoint
set ruleBeforeEnd to end of content of text object of boundDoc
set ruleBeforeParagraphs to count paragraphs of boundDoc
set rulePrefix to ""
if insertionPoint > 0 then
set rulePrefixRange to create range boundDoc start 0 end insertionPoint
set rulePrefix to content of rulePrefixRange as text
end if
set ruleRange to text object of paragraph ruleBeforeParagraphs of boundDoc
set ruleStart to start of content of ruleRange
set alignment of paragraph format of ruleRange to align paragraph center
set ownBorder to get border ruleRange which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set color of ownBorder to {49344,49344,49344}
set insertionRange to create range boundDoc start ruleOriginalPoint end ruleOriginalPoint
set content of insertionRange to " " & return
set ruleRange to text object of paragraph ruleBeforeParagraphs of boundDoc
set ownBorder to get border ruleRange which border border bottom
set insertedRange to create range boundDoc start ruleOriginalPoint end (ruleOriginalPoint + 2)
set followingRange to create range boundDoc start (ruleOriginalPoint + 2) end (ruleOriginalPoint + 3)
set followingStyle to name local of style of followingRange as text
set followingFormat to paragraph format of followingRange
set followingBefore to space before of followingFormat
set followingAfter to space after of followingFormat
set followingLeft to paragraph format left indent of followingFormat
set followingRight to paragraph format right indent of followingFormat
set followingFirst to first line indent of followingFormat
set followingSpacing to line spacing of followingFormat
set followingSpacingRule to line spacing rule of followingFormat
set followingBorder to get border followingRange which border border bottom
set line style of followingBorder to line style none
set followingFormat to paragraph format of followingRange
set followingPreserved to ((name local of style of followingRange as text) is followingStyle and space before of followingFormat is followingBefore and space after of followingFormat is followingAfter and paragraph format left indent of followingFormat is followingLeft and paragraph format right indent of followingFormat is followingRight and first line indent of followingFormat is followingFirst and line spacing of followingFormat is followingSpacing and line spacing rule of followingFormat is followingSpacingRule and (alignment of followingFormat is align paragraph center))
set prefixPreserved to true
if ruleOriginalPoint > 0 then
set rulePrefixRange to create range boundDoc start 0 end ruleOriginalPoint
set prefixPreserved to ((current application's NSString's stringWithString:(content of rulePrefixRange as text))'s isEqualToString:rulePrefix) as boolean
end if
set nativeRows to {{"paragraph-rule",ruleOriginalPoint,ruleStart,end of content of ruleRange,ruleBeforeEnd,end of content of text object of boundDoc,ruleBeforeParagraphs,count paragraphs of boundDoc,prefixPreserved,content of insertedRange as text,content of followingRange as text,(alignment of paragraph format of ruleRange is align paragraph center),(line style of ownBorder is line style single),(line width of ownBorder is line width75 point),(color of ownBorder is {49344,49344,49344}),(line style of followingBorder is line style none),followingPreserved}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
