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
if (count documents) is not 0 then error "WPSC_DIRECT_OPEN_REQUIRES_EMPTY"
set openedValue to open (POSIX file "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-mno3rsn5/direct-open-4c7bb28fca1249a98c2d1e0e5470a6cd.xml") file converter open format xmldocument serialized read only true add to recent files false confirm conversions false
set resultKind to "missing value"
set resultCount to 0
if openedValue is not missing value then
set resultKind to (class of openedValue) as text
set resultCount to 1
if class of openedValue is list then set resultCount to count openedValue
end if
set nativeRows to {{"open-result",resultKind,resultCount}}
repeat with nativeDocument in documents
set end of nativeRows to {"document",name of nativeDocument as text,posix full name of nativeDocument as text,saved of nativeDocument,id of active window of nativeDocument}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
