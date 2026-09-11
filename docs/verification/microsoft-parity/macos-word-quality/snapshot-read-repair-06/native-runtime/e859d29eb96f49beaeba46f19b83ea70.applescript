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
set boundDoc to document "document-1ae4884692704bb39a15dc26a88979a4.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-xz62ytzb/document-1ae4884692704bb39a15dc26a88979a4.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "main-text-original-start:0: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "main-text-original-start:1: set probeStoryRange to get story range boundDoc story type main text story"
log probePhase
set probeStoryRange to get story range boundDoc story type main text story
set probePhase to "main-text-original-start:2: set probeValue to \"unresolved-sentinel\""
log probePhase
set probeValue to "unresolved-sentinel"
set probePhase to "main-text-original-start:3: set probeValue to get start of content of probeStoryRange"
log probePhase
set probeValue to get start of content of probeStoryRange
set probePhase to "main-text-original-start:4: set probeValueClass to (class of probeValue) as text"
log probePhase
set probeValueClass to (class of probeValue) as text
set probePhase to "main-text-original-start:5: set probeRows to {{\"scalar\",probeValueClass,probeValue as integer}}"
log probePhase
set probeRows to {{"scalar",probeValueClass,probeValue as integer}}
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
