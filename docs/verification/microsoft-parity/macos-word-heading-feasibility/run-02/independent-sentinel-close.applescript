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
set sentinelDocument to document "文档70"
if (name of sentinelDocument as text) is not "文档70" then error "SENTINEL_IDENTITY_CHANGED"
if (path of sentinelDocument as text) is not "" then error "SENTINEL_SAVED"
if (content of text object of sentinelDocument as text) is not "HEADING FEASIBILITY SENTINEL 3794517177ea40f8868b09655d51bd1e 中文😀" & return then error "SENTINEL_CHANGED"
if saved of sentinelDocument then error "SENTINEL_SAVED"
close sentinelDocument saving no
set nativeRows to {{"sentinel-closed","文档70"}}
end tell
return my jsonRows(nativeRows)