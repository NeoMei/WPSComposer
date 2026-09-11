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
set boundDoc to document "document-810cc00a5e4b4af3961bc3e339bab8c6.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-ch978a__/document-810cc00a5e4b4af3961bc3e339bab8c6.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {{"body",content of text object of boundDoc as text}}
repeat with captionIndex from 1 to count fields of boundDoc
set captionField to field captionIndex of boundDoc
set end of nativeRows to {"field",content of field code of captionField as text,content of result range of captionField as text}
end repeat
set captionBookmark to bookmark "wpsc_fig_cccccccccccccccccccccccc" of boundDoc
set end of nativeRows to {"bookmark",name of captionBookmark,start of bookmark of captionBookmark,end of bookmark of captionBookmark,content of text object of captionBookmark as text}
set captionBookmark to bookmark "wpsc_tab_dddddddddddddddddddddddd" of boundDoc
set end of nativeRows to {"bookmark",name of captionBookmark,start of bookmark of captionBookmark,end of bookmark of captionBookmark,content of text object of captionBookmark as text}
repeat with captionIndex from 1 to count paragraphs of boundDoc
set captionRange to text object of paragraph captionIndex of boundDoc
set captionText to content of captionRange as text
if captionText contains "图😀(" or captionText contains "表😀[" then
set end of nativeRows to {"paragraph",start of content of captionRange,end of content of captionRange,captionText,(alignment of paragraph format of captionRange is align paragraph center),keep together of paragraph format of captionRange,keep with next of paragraph format of captionRange}
end if
end repeat
set end of nativeRows to {"counts",count tables of boundDoc,count inline shapes of boundDoc,count shapes of boundDoc}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
