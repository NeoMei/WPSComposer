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
set boundDoc to document "document-37330cef3146407e8a0705181240ae0d.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-n1dpth4d/document-37330cef3146407e8a0705181240ae0d.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {}
set bodyRange to text object of boundDoc
set end of nativeRows to {"document",content of bodyRange as text,(end of content of bodyRange) - 1,end of content of bodyRange,count characters of bodyRange,count paragraphs of boundDoc,count tables of boundDoc,count fields of boundDoc}
repeat with pi from 1 to count paragraphs of boundDoc
set ownRange to text object of paragraph pi of boundDoc
set end of nativeRows to {"paragraph",pi,content of ownRange as text,start of content of ownRange,end of content of ownRange}
end repeat
repeat with ti from 1 to count tables of boundDoc
set ownTable to table ti of boundDoc
set end of nativeRows to {"table",ti,start of content of text object of ownTable,end of content of text object of ownTable,count rows of ownTable,count columns of ownTable}
end repeat
repeat with bi from 1 to count bookmarks of boundDoc
set ownBookmark to bookmark bi of boundDoc
set end of nativeRows to {"bookmark",bi,name of ownBookmark as text,content of text object of ownBookmark as text,start of content of text object of ownBookmark,end of content of text object of ownBookmark}
end repeat
repeat with fi from 1 to count fields of boundDoc
set ownField to field fi of boundDoc
set end of nativeRows to {"field",fi,content of field code of ownField as text,content of result range of ownField as text,start of content of result range of ownField,end of content of result range of ownField}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
