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
set sourceStyle to Word style (style heading1) of boundDoc
set detachedStyle to Word style "WPSC Heading Clone" of boundDoc
set sourceFont to font object of sourceStyle
set destinationFont to font object of detachedStyle
set sourceParagraph to paragraph format of sourceStyle
set destinationParagraph to paragraph format of detachedStyle
set observedStyleName to name local of style of text object of paragraph 2 of boundDoc as text
if not (((current application's NSString's stringWithString:observedStyleName)'s isEqualToString:"WPSC Heading Clone") as boolean) then error "WPSC_FORMATTED_STYLE_IDENTITY"
set sourceStart to start of content of text object of paragraph 1 of boundDoc
set sourceEnd to (end of content of text object of paragraph 1 of boundDoc) - 1
set targetStart to start of content of text object of paragraph 2 of boundDoc
set targetEnd to (end of content of text object of paragraph 2 of boundDoc) - 1
set sourceTextRange to create range boundDoc start sourceStart end sourceEnd
set targetTextRange to create range boundDoc start targetStart end targetEnd
set formatted text of targetTextRange to formatted text of sourceTextRange
set observedStyleName to name local of style of text object of paragraph 2 of boundDoc as text
if not (((current application's NSString's stringWithString:observedStyleName)'s isEqualToString:"WPSC Heading Clone") as boolean) then error "WPSC_FORMATTED_STYLE_IDENTITY"
set paragraph format of detachedStyle to paragraph format of sourceStyle
set expectedBase to name local of Word style (style normal) of boundDoc as text
set observedBase to name local of Word style (base style of detachedStyle) of boundDoc as text
if not (((current application's NSString's stringWithString:observedBase)'s isEqualToString:expectedBase) as boolean) then error "WPSC_FORMATTED_BASE_CHANGED"
set automatically update of detachedStyle to true
activate object boundWindow
set selection start of selection of boundWindow to targetStart
set selection end of selection of boundWindow to targetEnd
if (selection start of selection of boundWindow is not targetStart) or (selection end of selection of boundWindow is not targetEnd) then error "WPSC_FORMATTED_SELECTION_CHANGED"
set observedStyleName to name local of style of selection of boundWindow as text
if not (((current application's NSString's stringWithString:observedStyleName)'s isEqualToString:"WPSC Heading Clone") as boolean) then error "WPSC_FORMATTED_STYLE_IDENTITY"
set style of selection of boundWindow to detachedStyle
set automatically update of detachedStyle to false
set observedStyleName to name local of style of text object of paragraph 2 of boundDoc as text
if not (((current application's NSString's stringWithString:observedStyleName)'s isEqualToString:"WPSC Heading Clone") as boolean) then error "WPSC_FORMATTED_STYLE_IDENTITY"
if automatically update of detachedStyle then error "WPSC_FORMATTED_STYLE_AUTOUPDATE"
set nativeRows to {{"stage",true}}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
