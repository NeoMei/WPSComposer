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
set boundDoc to document "document-56462579dcb84c729ce1a2b6c7b5bb50.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-fyj72s8k/document-56462579dcb84c729ce1a2b6c7b5bb50.docx" then error "WPSC_STALE_DOCUMENT"
set paginationBookmark to bookmark "WPSC_Pagination" of boundDoc
set paginationRange to text object of paragraph 1 of text object of paginationBookmark
set paginationPage to get range information paginationRange information type active end page number
set paginationX to get range information paginationRange information type horizontal position relative to page
set paginationY to get range information paginationRange information type vertical position relative to page
set nativeRows to {{"bookmark",start of content of paginationRange,end of content of paginationRange,paginationPage,paginationX,paginationY}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
