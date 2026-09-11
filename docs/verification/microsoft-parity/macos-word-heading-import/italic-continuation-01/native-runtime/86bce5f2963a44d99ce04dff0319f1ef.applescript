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
set boundDoc to document "document-bbcdd06a9c974e52bb2b0d3f0da20bbb.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-txw2p1e4/document-bbcdd06a9c974e52bb2b0d3f0da20bbb.docx" then error "WPSC_STALE_DOCUMENT"
if not (((current application's NSString's stringWithString:(content of text object of boundDoc as text))'s isEqualToString:"FRESH STYLE 中文 العربية 😀" & return & "IMPORTED APPEARANCE 中文 العربية 😀" & return & "SUFFIX" & return & "FRESH STYLE 中文 العربية 😀" & return & "") as boolean) then error "WPSC_ITALIC_BODY"
set sourceStyle to Word style (style heading1) of boundDoc
set linkedStyle to Word style "标题 1 字符" of boundDoc
if italic of font object of sourceStyle is not false then error "WPSC_ITALIC_PREIMAGE"
if italic of font object of linkedStyle is not false then error "WPSC_ITALIC_PAIR_PREIMAGE"
set italic of font object of sourceStyle to true
set nativeRows to {{"italic",italic of font object of sourceStyle,italic of font object of linkedStyle}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
