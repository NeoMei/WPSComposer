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
try
set tabsValue to tab stops of paragraph 1 of r
set end of nativeRows to {"rangepara","count",count tab stops of paragraph 1 of r}
repeat with t in tab stops of paragraph 1 of r
set end of nativeRows to {"rangepara",tab stop position of t}
end repeat
on error msg number num
set end of nativeRows to {"rangepara","error",num,msg}
end try
try
set tabsValue to tab stops of p
set end of nativeRows to {"para","count",count tab stops of p}
repeat with t in tab stops of p
set end of nativeRows to {"para",tab stop position of t}
end repeat
on error msg number num
set end of nativeRows to {"para","error",num,msg}
end try
try
set tabsValue to tab stops of paragraph format of r
set end of nativeRows to {"format","count",count tab stops of paragraph format of r}
repeat with t in tab stops of paragraph format of r
set end of nativeRows to {"format",tab stop position of t}
end repeat
on error msg number num
set end of nativeRows to {"format","error",num,msg}
end try
try
set tabsValue to tab stops of paragraph ((count paragraphs of d) - 1) of d
set end of nativeRows to {"direct","count",count tab stops of paragraph ((count paragraphs of d) - 1) of d}
repeat with t in tab stops of paragraph ((count paragraphs of d) - 1) of d
set end of nativeRows to {"direct",tab stop position of t}
end repeat
on error msg number num
set end of nativeRows to {"direct","error",num,msg}
end try
repeat with di from 1 to count documents
set inventoryDoc to document di
set documentText to content of text object of inventoryDoc as text
set textHash to do shell script ("/usr/bin/printf %s " & quoted form of documentText & " | /usr/bin/shasum -a 256")
set end of nativeRows to {"inventory",name of inventoryDoc as text,posix full name of inventoryDoc as text,saved of inventoryDoc,textHash}
end repeat
end tell
return my jsonRows(nativeRows)