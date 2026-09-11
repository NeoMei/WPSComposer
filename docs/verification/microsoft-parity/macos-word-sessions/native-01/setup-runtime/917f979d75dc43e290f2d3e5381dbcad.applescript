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
set rows to {}
set seedDoc to make new document
set content of text object of seedDoc to "Alpha" & return & "Beta" & return & "Gamma" & return
save as seedDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-6nvsgqcz/seed.docx" file format format document default add to recent files false
close seedDoc saving no
set sentinelDoc to make new document
set content of text object of sentinelDoc to "WPSC-SENTINEL-a0c5890252e54b62a24af1d158fb7a53"
set rows to {{"binding", id of active window of sentinelDoc, posix full name of sentinelDoc as text, name of sentinelDoc as text}}
end tell
end timeout
return my jsonRows(rows)
