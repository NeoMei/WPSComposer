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
set boundDoc to document "document-504defba49bc4b74ad332f3d39b24dae.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jfh1jacl/document-504defba49bc4b74ad332f3d39b24dae.docx" then error "WPSC_STALE_DOCUMENT"
set captionDetail to create range boundDoc start 13 end 67
set captionMark to create range boundDoc start 67 end 68
set followingDetail to create range boundDoc start 68 end 69
set nativeRows to {{"caption-format-detail","global",start of content of captionDetail,end of content of captionDetail,content of captionDetail as text,(alignment of paragraph format of captionDetail is align paragraph center),keep together of paragraph format of captionDetail,keep with next of paragraph format of captionDetail,first line indent of paragraph format of captionDetail,paragraph format left indent of paragraph format of captionDetail,paragraph format right indent of paragraph format of captionDetail,space before of paragraph format of captionDetail,space after of paragraph format of captionDetail,line spacing of paragraph format of captionDetail,italic of font object of captionDetail,font size of font object of captionDetail},{"caption-mark-detail","global",start of content of captionMark,end of content of captionMark,content of captionMark as text},{"following-text-detail","global",start of content of followingDetail,end of content of followingDetail,content of followingDetail as text}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
