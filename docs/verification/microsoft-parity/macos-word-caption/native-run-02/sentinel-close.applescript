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

tell application "Microsoft Word"
set qualitySentinel to document "文档27"
if not ((current application's NSString's stringWithString:(name of qualitySentinel as text))'s isEqualToString:"文档27") then error "QUALITY_SENTINEL_NAME"
if not ((current application's NSString's stringWithString:(path of qualitySentinel as text))'s isEqualToString:"") then error "QUALITY_SENTINEL_SAVED"
if not ((current application's NSString's stringWithString:(content of text object of qualitySentinel as text))'s isEqualToString:"CAPTION SENTINEL 中文😀 ce54641d046b474c83a693bbea2b02de" & return & "") then error "QUALITY_SENTINEL_TEXT"
if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"
close qualitySentinel saving no
set nativeRows to {{"sentinel-closed"}}
end tell
return my jsonRows(nativeRows)
