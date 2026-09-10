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
set boundDoc to document "document-22e1a0124e2d4a8a8c21c41e4e45f824.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-myaybcl6/document-22e1a0124e2d4a8a8c21c41e4e45f824.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set content of text object of boundDoc to "HEAD" & return & "left REPLACE-长😀 right" & return & "TAIL" & return & ""
set probeTarget to create range boundDoc start 10 end 21
make new bookmark at boundDoc with properties {name:"figure_probe_target",text object:probeTarget}
set probeSuffix to create range boundDoc start 21 end 27
make new bookmark at boundDoc with properties {name:"figure_probe_suffix",text object:probeSuffix}
set first line indent of paragraph format of probeTarget to 17
set paragraph format left indent of paragraph format of probeTarget to 3
set paragraph format right indent of paragraph format of probeTarget to 5
set space before of paragraph format of probeTarget to 11
set space after of paragraph format of probeTarget to 7
set line spacing rule of paragraph format of probeTarget to line space exactly
set line spacing of paragraph format of probeTarget to 21
set italic of font object of probeTarget to true
set font size of font object of probeTarget to 14
set alignment of paragraph format of probeTarget to align paragraph right
set keep together of paragraph format of probeTarget to false
set keep with next of paragraph format of probeTarget to false
set selection start of selection of boundWindow to 10
set selection end of selection of boundWindow to 21
set nativeRows to {{"seed",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
