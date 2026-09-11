use framework "Foundation"
use scripting additions
on jsonRows(rows)
  set dataValue to current application's NSJSONSerialization's dataWithJSONObject:rows options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:dataValue encoding:4) as text
end jsonRows
if not application "Microsoft Word" is running then error "NOT_RUNNING"
tell application "Microsoft Word"
set nativeRows to {}
set d to document "document-e099edb5171c4eca8b004158045dc5e3.docx"
if (posix full name of d as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-pdx7rs5d/document-e099edb5171c4eca8b004158045dc5e3.docx" then error "WRONG_DOCUMENT"
set docEnd to end of content of text object of d
set end of nativeRows to {"end",docEnd}
set r to create range d start (docEnd - 8) end docEnd
set end of nativeRows to {"tail",content of r as text}
set r to create range d start (docEnd - 1) end (docEnd - 1)
set rawValue to content of r
set end of nativeRows to {"empty",start of content of r,end of content of r,class of rawValue as text,rawValue as text,count characters of (rawValue as text)}
repeat with fi from 1 to count fields of d
set f to field fi of d
set end of nativeRows to {"field",fi,content of field code of f as text,start of content of field code of f,start of content of result range of f,end of content of result range of f,content of result range of f as text}
end repeat
repeat with di from 1 to count documents
set inventoryDoc to document di
set documentText to content of text object of inventoryDoc as text
set textHash to do shell script ("/usr/bin/printf %s " & quoted form of documentText & " | /usr/bin/shasum -a 256")
set end of nativeRows to {"inventory",name of inventoryDoc as text,posix full name of inventoryDoc as text,saved of inventoryDoc,textHash}
end repeat
end tell
return my jsonRows(nativeRows)
