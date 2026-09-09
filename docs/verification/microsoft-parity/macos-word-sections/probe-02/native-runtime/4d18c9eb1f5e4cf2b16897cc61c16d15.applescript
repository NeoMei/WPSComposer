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
set boundDoc to document "document-8339d0eb0e2d4014a82ef86cd16754f7.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-vcggy6oq/document-8339d0eb0e2d4014a82ef86cd16754f7.docx" then error "WPSC_STALE_DOCUMENT"
set content of text object of boundDoc to "Probe body 中文😀" & return
make new variable at boundDoc with properties {name:"WpsComposerSectionRole_1", variable value:"封面😀"}
set variable value of variable "WpsComposerSectionRole_1" of boundDoc to "正文😀"
set value of document property "Title" of boundDoc to "标题😀"
set value of document property "Author" of boundDoc to "作者😀"
set ownFooter to get footer (section 1 of boundDoc) index header footer primary
create new field text range (text object of ownFooter) field type field page preserve formatting true
insert text "Retain footer " at beginning of text object of ownFooter
set number style of page number options of ownFooter to page number style lowercase roman
set restart numbering at section of page number options of ownFooter to true
set starting number of page number options of ownFooter to 3
set ownFooter to get footer (section 1 of boundDoc) index header footer primary
set nativeRows to {{"version", version}, {"role", variable value of variable "WpsComposerSectionRole_1" of boundDoc}, {"title", value of document property "Title" of boundDoc}, {"author", value of document property "Author" of boundDoc}, {"footer", content of text object of ownFooter}, {"number", number style of page number options of ownFooter as text, restart numbering at section of page number options of ownFooter, starting number of page number options of ownFooter}}
set footerRange to text object of ownFooter
repeat with fieldIndex from 1 to (count fields of footerRange)
set ownField to field fieldIndex of footerRange
set end of nativeRows to {"field", field type of ownField as text, content of field code of ownField as text}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
