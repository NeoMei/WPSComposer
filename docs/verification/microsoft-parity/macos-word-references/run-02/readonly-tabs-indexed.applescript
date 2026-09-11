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
if not application "Microsoft Word" is running then error "NOT_RUNNING"
tell application "Microsoft Word"
set nativeRows to {}
set d to document "document-6fe260f50f6e4df181b7fcbc16f27abc.docx"
if (posix full name of d as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-eddiheun/document-6fe260f50f6e4df181b7fcbc16f27abc.docx" then error "WRONG_DOCUMENT"
set p to paragraph ((count paragraphs of d) - 1) of d
set r to text object of p
set nativeRows to {{"paragraph",start of content of r,end of content of r,content of r as text}}
repeat with ti from 1 to count tab stops of paragraph 1 of r
set t to tab stop ti of paragraph 1 of r
set end of nativeRows to {"indexed-tab",ti,tab stop position of t,custom tab of t}
end repeat
repeat with di from 1 to count documents
set inventoryDoc to document di
set documentText to content of text object of inventoryDoc as text
set textHash to do shell script ("/usr/bin/printf %s " & quoted form of documentText & " | /usr/bin/shasum -a 256")
set end of nativeRows to {"inventory",name of inventoryDoc as text,posix full name of inventoryDoc as text,saved of inventoryDoc,textHash}
end repeat
end tell
return my jsonRows(nativeRows)