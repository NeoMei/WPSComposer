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
set boundDoc to document "document-62b78f798da34e6285d912d125996c1b.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-owklc5il/document-62b78f798da34e6285d912d125996c1b.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {}
set r to create range boundDoc start 382 end 447
set tabVerified to false
repeat with tabIndex from 1 to count tab stops of paragraph 1 of r
set ownTab to tab stop tabIndex of paragraph 1 of r
if tab stop position of ownTab is 24 then set tabVerified to true
end repeat
set end of nativeRows to {"list",paragraph format left indent of paragraph format of r,first line indent of paragraph format of r,(line spacing rule of paragraph format of r is line space1 pt5),space before of paragraph format of r,space after of paragraph format of r,tabVerified}
set r to create range boundDoc start 447 end 472
set tabVerified to false
repeat with tabIndex from 1 to count tab stops of paragraph 1 of r
set ownTab to tab stop tabIndex of paragraph 1 of r
if tab stop position of ownTab is 31.5 then set tabVerified to true
end repeat
set end of nativeRows to {"list",paragraph format left indent of paragraph format of r,first line indent of paragraph format of r,(line spacing rule of paragraph format of r is line space1 pt5),space before of paragraph format of r,space after of paragraph format of r,tabVerified}
set r to create range boundDoc start 453 end 462
set end of nativeRows to {"degradation",start of content of r,end of content of r,content of r as text,italic of font object of r,(color of font object of r is {40092, 0, 1542}),(background pattern color of shading of r is {64764, 59624, 59110})}
set r to create range boundDoc start 462 end 471
set end of nativeRows to {"degradation",start of content of r,end of content of r,content of r as text,italic of font object of r,(color of font object of r is {40092, 0, 1542}),(background pattern color of shading of r is {64764, 59624, 59110})}
set r to create range boundDoc start 472 end 480
set end of nativeRows to {"static-style",italic of font object of r,(color of font object of r is {40092, 0, 1542}),(background pattern color of shading of r is {64764, 59624, 59110})}
set r to create range boundDoc start 480 end 493
set end of nativeRows to {"static-style",italic of font object of r,(color of font object of r is {40092, 0, 1542}),(background pattern color of shading of r is {64764, 59624, 59110})}
set r to create range boundDoc start 493 end 494
set end of nativeRows to {"static-style",italic of font object of r,(color of font object of r is {40092, 0, 1542}),(background pattern color of shading of r is {64764, 59624, 59110})}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
