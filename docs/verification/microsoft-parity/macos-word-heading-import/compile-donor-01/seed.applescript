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
activate object boundWindow
set sourceStyle to Word style (style heading1) of boundDoc
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
set content of text object of boundDoc to "SOURCE APPEARANCE 中文 العربية 😀" & return & "REPLACE" & return & "SUFFIX" & return & ""
set style of text object of paragraph 1 of boundDoc to sourceStyle
set nativeRows to {{"stage",true}}
end tell
return my jsonRows(nativeRows)
