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
set nativeRows to {}
if application "Microsoft Word" is running then
tell application "Microsoft Word"
try
set observation to exists document "document-210ca8985c0942949d66360c6de56031.docx"
set end of nativeRows to {"private-exists",observation as text}
on error errorText number errorNumber
set end of nativeRows to {"private-exists",errorNumber,errorText}
end try
try
set observation to name of document "document-210ca8985c0942949d66360c6de56031.docx" as text
set end of nativeRows to {"private-name",observation as text}
on error errorText number errorNumber
set end of nativeRows to {"private-name",errorNumber,errorText}
end try
try
set observation to properties of document "document-210ca8985c0942949d66360c6de56031.docx"
set end of nativeRows to {"private-properties",observation as text}
on error errorText number errorNumber
set end of nativeRows to {"private-properties",errorNumber,errorText}
end try
try
set observation to name of active document as text
set end of nativeRows to {"active-doc-name",observation as text}
on error errorText number errorNumber
set end of nativeRows to {"active-doc-name",errorNumber,errorText}
end try
try
set observation to name of active window as text
set end of nativeRows to {"active-window-name",observation as text}
on error errorText number errorNumber
set end of nativeRows to {"active-window-name",errorNumber,errorText}
end try
try
set observation to every document
set end of nativeRows to {"all-documents",observation as text}
on error errorText number errorNumber
set end of nativeRows to {"all-documents",errorNumber,errorText}
end try
try
set observation to every window
set end of nativeRows to {"all-windows",observation as text}
on error errorText number errorNumber
set end of nativeRows to {"all-windows",errorNumber,errorText}
end try
end tell
end if
return my jsonRows(nativeRows)