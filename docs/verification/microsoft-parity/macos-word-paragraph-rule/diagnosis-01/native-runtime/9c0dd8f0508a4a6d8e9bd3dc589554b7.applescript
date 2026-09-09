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
set boundDoc to document "document-9c5718eabc104877b1efb05645d14d8e.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-o3ck5g8k/document-9c5718eabc104877b1efb05645d14d8e.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {{"stage","after_public_rule"}}
repeat with paragraphIndex from 1 to count paragraphs of boundDoc
set captureRange to text object of paragraph paragraphIndex of boundDoc
try
set captureValue to (content of captureRange as text)
set end of nativeRows to {paragraphIndex,"content",captureValue}
on error detail number errorNumber
set end of nativeRows to {paragraphIndex,"content","ERROR",errorNumber,detail}
end try
try
set captureValue to (style of captureRange is Word style (style body text) of boundDoc)
set end of nativeRows to {paragraphIndex,"style_equals_object",captureValue}
on error detail number errorNumber
set end of nativeRows to {paragraphIndex,"style_equals_object","ERROR",errorNumber,detail}
end try
try
set captureValue to (style of captureRange as text)
set end of nativeRows to {paragraphIndex,"style_as_text",captureValue}
on error detail number errorNumber
set end of nativeRows to {paragraphIndex,"style_as_text","ERROR",errorNumber,detail}
end try
try
set captureValue to (name local of style of captureRange as text)
set end of nativeRows to {paragraphIndex,"style_name_local",captureValue}
on error detail number errorNumber
set end of nativeRows to {paragraphIndex,"style_name_local","ERROR",errorNumber,detail}
end try
try
set captureValue to (name local of Word style (style body text) of boundDoc as text)
set end of nativeRows to {paragraphIndex,"bodytext_name_local",captureValue}
on error detail number errorNumber
set end of nativeRows to {paragraphIndex,"bodytext_name_local","ERROR",errorNumber,detail}
end try
try
set captureValue to (Word style (style body text) of boundDoc as text)
set end of nativeRows to {paragraphIndex,"bodytext_as_text",captureValue}
on error detail number errorNumber
set end of nativeRows to {paragraphIndex,"bodytext_as_text","ERROR",errorNumber,detail}
end try
try
set captureValue to (alignment of paragraph format of captureRange is align paragraph center)
set end of nativeRows to {paragraphIndex,"centered",captureValue}
on error detail number errorNumber
set end of nativeRows to {paragraphIndex,"centered","ERROR",errorNumber,detail}
end try
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
