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
if not application "Microsoft Word" is running then error "SENTINEL_UNAVAILABLE"
tell application "Microsoft Word"
set sentinelDocument to document "文档32"
if (name of sentinelDocument as text) is not "文档32" then error "SENTINEL_IDENTITY_CHANGED"
if (path of sentinelDocument as text) is not "" then error "SENTINEL_SAVED"
if (content of text object of sentinelDocument as text) is not "Notice sentinel 中文😀 668159b54d4246d9a4507d0f11dec00c" & return then error "SENTINEL_CHANGED"
if saved of sentinelDocument then error "SENTINEL_SAVED"
close sentinelDocument saving no
set nativeRows to {{"sentinel-closed","文档32"}}
end tell
return my jsonRows(nativeRows)