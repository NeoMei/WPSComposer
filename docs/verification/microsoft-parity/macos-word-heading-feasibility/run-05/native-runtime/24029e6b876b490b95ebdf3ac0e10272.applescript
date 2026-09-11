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
set boundDoc to document "document-ab1edbfbaa69478a8d9e770acca96f46.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-piet06hs/document-ab1edbfbaa69478a8d9e770acca96f46.docx" then error "WPSC_STALE_DOCUMENT"
activate object boundWindow
set sourceStyle to Word style (style heading1) of boundDoc
set detachedStyle to make new Word style at boundDoc with properties {name local:"WPSC Heading Clone"}
set base style of detachedStyle to style normal
set sourceFont to font object of sourceStyle
set sourceParagraph to paragraph format of sourceStyle
set name of sourceFont to "Arial"
set ascii name of sourceFont to "Arial"
set font size of sourceFont to 17
set bold of sourceFont to true
set italic of sourceFont to true
set scaling of sourceFont to 110
set spacing of sourceFont to 1.25
set kerning of sourceFont to 12
set color of sourceFont to {39936,0,1536}
set background pattern color of shading of sourceFont to {64512,59392,58880}
set outside line style of border options of sourceFont to line style single
set outside color of border options of sourceFont to {49344,49344,49344}
set alignment of sourceParagraph to align paragraph center
set space before of sourceParagraph to 14
set space after of sourceParagraph to 7
set first line indent of sourceParagraph to 4
set keep with next of sourceParagraph to true
set widow control of sourceParagraph to false
set line spacing rule of sourceParagraph to line space exactly
set line spacing of sourceParagraph to 23
set background pattern color of shading of sourceParagraph to {58880,62208,65280}
set outside line style of border options of sourceParagraph to line style single
set outside line width of border options of sourceParagraph to line width75 point
set outside color of border options of sourceParagraph to {32768,16384,8192}
make new tab stop at sourceParagraph with properties {tab stop position:40,alignment:align tab right}
set content of text object of boundDoc to "SOURCE APPEARANCE 中文😀" & return & "DETACHED APPEARANCE 中文😀" & return
set style of text object of paragraph 1 of boundDoc to sourceStyle
set style of text object of paragraph 2 of boundDoc to detachedStyle
set content of text object of boundDoc to "FORMATTED STYLE 中文 العربية" & return & "FORMATTED STYLE 中文 العربية" & return
set style of text object of paragraph 1 of boundDoc to sourceStyle
set style of text object of paragraph 2 of boundDoc to detachedStyle
set nativeRows to {{"stage",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
