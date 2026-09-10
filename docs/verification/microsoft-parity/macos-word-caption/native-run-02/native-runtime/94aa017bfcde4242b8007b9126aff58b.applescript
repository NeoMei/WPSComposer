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
set boundDoc to document "document-b1eb906af6ea4dab84556d30f7c90554.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-4djbbora/document-b1eb906af6ea4dab84556d30f7c90554.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set content of text object of boundDoc to "Chapter" & return & "left REPLACE-长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀长😀 right" & return & "" & return & "TAIL" & return & ""
set captionList to make new list template at boundDoc with properties {name:"WPSCCaptionFixture",outline numbered:true}
set captionLevel to list level 1 of captionList
set number style of captionLevel to list number style arabic
set number format of captionLevel to "%1"
set start at of captionLevel to 1
set linked style of captionLevel to name local of Word style (style heading1) of boundDoc as text
set style of text object of paragraph 1 of boundDoc to style heading1
set secondCaptionRange to create range boundDoc start 142 end 147
make new bookmark at boundDoc with properties {name:"caption_second_slot",text object:secondCaptionRange}
set selection start of selection of boundWindow to 13
set selection end of selection of boundWindow to 141
set nativeRows to {{"seed",list string of list format of text object of paragraph 1 of boundDoc as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
