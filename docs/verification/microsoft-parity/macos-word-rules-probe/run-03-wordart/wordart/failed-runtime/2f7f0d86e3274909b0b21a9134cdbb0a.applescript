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
set boundDoc to document "document-b969b5e05de94fcdb02f77b174ce8740.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-7ht6kz6c/document-b969b5e05de94fcdb02f77b174ce8740.docx" then error "WPSC_STALE_DOCUMENT"
set p to (end of content of text object of boundDoc) - 1
set insertionRange to create range boundDoc start p end p
set beforeCount to count shapes of boundDoc
make new word art at boundDoc with properties {anchor:insertionRange,word art text:"WORDART 中文😀",font name:"Arial",font size:36,bold:false,italic:false,left position:200,top:400,preset word art effect:wordart format1}
if (count shapes of boundDoc) is not beforeCount + 1 then error "WPSC_WORDART_COUNT_DELTA_FAILED"
set ownArt to shape (count shapes of boundDoc) of boundDoc
if (shape type of ownArt) is not shape type word art then error "WPSC_WORDART_NATIVE_TYPE_FAILED"
set ownArt to shape 1 of boundDoc
set artFormat to word art format of ownArt
set nativeRows to {{"wordart",count shapes of boundDoc,(shape type of ownArt is shape type word art),word art text of artFormat as text,font name of artFormat as text,font size of artFormat,bold of artFormat,italic of artFormat,left position of ownArt,top of ownArt,(preset word art effect of artFormat is wordart format1),width of ownArt,height of ownArt}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
