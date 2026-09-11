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
set boundDoc to document "document-504defba49bc4b74ad332f3d39b24dae.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jfh1jacl/document-504defba49bc4b74ad332f3d39b24dae.docx" then error "WPSC_STALE_DOCUMENT"
set secondCaptionRange to text object of bookmark "caption_second_slot" of boundDoc
set selection start of selection of boundWindow to start of content of secondCaptionRange
set selection end of selection of boundWindow to end of content of secondCaptionRange
set captionSelection to text object of selection of boundWindow
set nativeRows to {{"selection",start of content of captionSelection,end of content of captionSelection,content of captionSelection as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
