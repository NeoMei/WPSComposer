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
set boundDoc to document "COMPILE_ONLY.docx"
set boundWindow to active window of boundDoc
set ownSelection to selection of boundWindow
if story type of ownSelection is not main text story then error "WPSC_INSTALL_STORY"
set headingBookmark to bookmark "WPSC_InstallHeading" of boundDoc
set sourceStyle to Word style (style heading1) of boundDoc
set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,start of content of text object of ownSelection,end of content of text object of ownSelection,start of bookmark of headingBookmark,end of bookmark of headingBookmark,content of text object of headingBookmark as text,count fields of boundDoc,count tables of boundDoc,font size of font object of sourceStyle,paragraph format left indent of paragraph format of sourceStyle}}
end tell
return my jsonRows(nativeRows)
