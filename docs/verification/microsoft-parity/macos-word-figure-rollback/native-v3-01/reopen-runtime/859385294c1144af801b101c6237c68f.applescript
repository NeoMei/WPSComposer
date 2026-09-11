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
set boundDoc to document "document-23731b476253482bba9eb8ae19dda21d.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-q0kij2a3/document-23731b476253482bba9eb8ae19dda21d.docx" then error "WPSC_STALE_DOCUMENT"
set probeSelection to text object of selection of boundWindow
set probeTargetBookmark to bookmark "figure_probe_target" of boundDoc
set probeSuffixBookmark to bookmark "figure_probe_suffix" of boundDoc
set probeFormatRange to text object of probeTargetBookmark
set nativeRows to {}
set end of nativeRows to {"body",content of text object of boundDoc as text}
set end of nativeRows to {"selection",start of content of probeSelection,end of content of probeSelection,content of probeSelection as text}
set end of nativeRows to {"bookmark",name of probeTargetBookmark,start of bookmark of probeTargetBookmark,end of bookmark of probeTargetBookmark,content of text object of probeTargetBookmark as text}
set end of nativeRows to {"bookmark",name of probeSuffixBookmark,start of bookmark of probeSuffixBookmark,end of bookmark of probeSuffixBookmark,content of text object of probeSuffixBookmark as text}
set end of nativeRows to {"format",first line indent of paragraph format of probeFormatRange,paragraph format left indent of paragraph format of probeFormatRange,paragraph format right indent of paragraph format of probeFormatRange,space before of paragraph format of probeFormatRange,space after of paragraph format of probeFormatRange,line spacing of paragraph format of probeFormatRange,italic of font object of probeFormatRange,font size of font object of probeFormatRange,(alignment of paragraph format of probeFormatRange is align paragraph right),keep together of paragraph format of probeFormatRange,keep with next of paragraph format of probeFormatRange}
set end of nativeRows to {"counts",count inline pictures of boundDoc,count shapes of boundDoc,count tables of boundDoc,count fields of boundDoc,count bookmarks of boundDoc}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
