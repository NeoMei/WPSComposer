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
set boundDoc to document "document-651fc661d3a64ce6a966ffc90e2e5750.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-hhk9noow/document-651fc661d3a64ce6a966ffc90e2e5750.docx" then error "WPSC_STALE_DOCUMENT"
set numberSelection to selection of boundWindow
if story type of numberSelection is not main text story then error "WPSC_NUMBERING_WRONG_STORY"
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-hhk9noow/document-651fc661d3a64ce6a966ffc90e2e5750.docx") then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-hhk9noow/document-651fc661d3a64ce6a966ffc90e2e5750.docx") then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of numberSelection as text))'s isEqualToString:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-hhk9noow/document-651fc661d3a64ce6a966ffc90e2e5750.docx") then error "WPSC_STALE_DOCUMENT"
set numberSelectionRange to text object of numberSelection
if (start of content of numberSelectionRange) is not 64 or (end of content of numberSelectionRange) is not 64 then error "WPSC_NUMBERING_SELECTION_CHANGED"
set numberRange to create range boundDoc start 20 end 64
set alignment of paragraph format of numberRange to align paragraph right
set keep together of paragraph format of numberRange to true
set nativeRows to {{"format",(alignment of paragraph format of numberRange is align paragraph right),keep together of paragraph format of numberRange}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
