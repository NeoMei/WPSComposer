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
set boundDoc to document "document-776550c623a24a3fa7beed008376159f.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-_4wuhnyv/document-776550c623a24a3fa7beed008376159f.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {}
repeat with si from 1 to count sections of boundDoc
set ownSection to section si of boundDoc
set ownFooter to get footer ownSection index header footer primary
set ownHeader to get header ownSection index header footer primary
set end of nativeRows to {page width of page setup of ownSection,page height of page setup of ownSection,orientation of page setup of ownSection as text,top margin of page setup of ownSection,bottom margin of page setup of ownSection,left margin of page setup of ownSection,right margin of page setup of ownSection,link to previous of ownHeader,link to previous of ownFooter,restart numbering at section of page number options of ownFooter,starting number of page number options of ownFooter}
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
