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
set boundDoc to document "document-4790ad340a744d19b79c18614d85987b.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-rnhywxiy/document-4790ad340a744d19b79c18614d85987b.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {{"body",content of text object of boundDoc as text}}
repeat with ni from 1 to count fields of boundDoc
set nf to field ni of boundDoc
set end of nativeRows to {"field",content of field code of nf as text,content of result range of nf as text}
end repeat
set nb to bookmark "wpsc_eq_aaaaaaaaaaaaaaaaaaaaaaaa" of boundDoc
set end of nativeRows to {"bookmark",name of nb,start of bookmark of nb,end of bookmark of nb,content of text object of nb as text}
set nb to bookmark "wpsc_eq_bbbbbbbbbbbbbbbbbbbbbbbb" of boundDoc
set end of nativeRows to {"bookmark",name of nb,start of bookmark of nb,end of bookmark of nb,content of text object of nb as text}
set nb to bookmark "wpsc_eq_cccccccccccccccccccccccc" of boundDoc
set end of nativeRows to {"bookmark",name of nb,start of bookmark of nb,end of bookmark of nb,content of text object of nb as text}
set end of nativeRows to {"counts",count tables of boundDoc,count inline shapes of boundDoc,count shapes of boundDoc}
repeat with ni from 1 to count paragraphs of boundDoc
set nr to text object of paragraph ni of boundDoc
if (content of nr as text) contains tab then set end of nativeRows to {"paragraph",content of nr as text,(alignment of paragraph format of nr is align paragraph right),keep together of paragraph format of nr}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
