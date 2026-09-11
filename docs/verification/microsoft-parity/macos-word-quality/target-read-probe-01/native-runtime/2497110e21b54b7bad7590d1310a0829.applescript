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
set boundDoc to document "document-2b532297486e4b5789d4a83f9e6f8ad4.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jeypd1s4/document-2b532297486e4b5789d4a83f9e6f8ad4.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jeypd1s4/document-2b532297486e4b5789d4a83f9e6f8ad4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jeypd1s4/document-2b532297486e4b5789d4a83f9e6f8ad4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jeypd1s4/document-2b532297486e4b5789d4a83f9e6f8ad4.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set probePhase to "point"
set nativeRows to {}
try
set probeLower to qualityPoint - 1
set probeUpper to qualityPoint + 1
set end of nativeRows to {"bounds",qualityPoint,probeLower,probeUpper}
repeat with probeOrdinal from 1 to count tables of boundDoc
set qt to table probeOrdinal of boundDoc
set probePhase to "compound-comparison"
set probeHit to false
if (start of content of text object of qt) <= qualityPoint + 1 and (end of content of text object of qt) >= qualityPoint - 1 then set probeHit to true
set end of nativeRows to {"compound",probeOrdinal as integer,probeHit}
end repeat
set end of nativeRows to {"done"}
on error probeError number probeNumber
if probeNumber is not -2763 then error probeError number probeNumber
set end of nativeRows to {"expected-read-error",probePhase,probeNumber}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
