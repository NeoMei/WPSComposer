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
set sentinelDoc to document "\u6587\u6863104"
if (posix full name of sentinelDoc as text) is not "\u6587\u6863104" then error "WPSC_SENTINEL_PATH_CHANGED"
if (content of text object of sentinelDoc as text) is not "WPSC-OBJECT-SENTINEL-3aa1593038d147fbaac1f868d350952d\r" then error "WPSC_SENTINEL_TEXT_CHANGED"
if saved of sentinelDoc then error "WPSC_SENTINEL_SAVED"
close sentinelDoc saving no
set nativeRows to {{"sentinel_closed"}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
