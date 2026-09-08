on textLength(valueText)
return (count of characters of valueText) as text
end textLength

on eventLine(capabilityName, stateName, detailText)
set AppleScript's text item delimiters to " "
set detailText to (paragraphs of (detailText as text)) as text
set AppleScript's text item delimiters to ""
log "EVENT" & tab & capabilityName & tab & stateName & tab & detailText
end eventLine

on run argv
set docxPath to item 1 of argv
set pdfPath to item 2 of argv
set inputPath to item 3 of argv
set inputText to read POSIX file inputPath as «class utf8»
set ownedDoc to missing value
set beforeDocs to {}
set coreOK to false
set snapshotReady to false
with timeout of 600 seconds
 tell application "Microsoft Word"
  my eventLine("engine", "observed", version)
  try
   my eventLine("snapshot_before", "attempted", "")
   repeat with documentIndex from 1 to (count of documents)
    set userDoc to document documentIndex
    set end of beforeDocs to {name of userDoc, posix full name of userDoc, saved of userDoc, content of text object of userDoc}
    my eventLine("existing_before", "observed", (name of userDoc) & "|" & (posix full name of userDoc) & "|saved=" & (saved of userDoc as text) & "|characters=" & (my textLength(content of text object of userDoc)))
   end repeat
   set snapshotReady to true
   my eventLine("snapshot_before", "succeeded", (count beforeDocs) as text)
   my eventLine("create", "attempted", "")
   set ownedDoc to make new document
   set ownedName to name of ownedDoc
   repeat with prior in beforeDocs
    if item 1 of prior is ownedName then
     set ownedDoc to missing value
     error "New document name collided with a preexisting document; refusing mutation"
    end if
   end repeat
   set ownedDoc to document ownedName
   my eventLine("owned_identity", "observed", ownedName)
   set content of text object of ownedDoc to inputText
   my eventLine("create", "succeeded", ownedName)

my eventLine("chinese_body", "attempted", "")
try
set bodyRange to text object of paragraph 4 of ownedDoc
set name of font object of bodyRange to "FangSong"
set east asian name of font object of bodyRange to "仿宋"
set font size of font object of bodyRange to 12
set first line indent of paragraph format of bodyRange to 24
set character unit first line indent of paragraph format of bodyRange to 2
my eventLine("chinese_body", "succeeded", "")
on error errorText number errorNumber
my eventLine("chinese_body", "failed", (errorNumber as text) & ": " & errorText)
end try

my eventLine("headings", "attempted", "")
try
set style of paragraph 3 of ownedDoc to style heading1
set style of paragraph 5 of ownedDoc to style heading2
set style of paragraph 7 of ownedDoc to style heading3
set style of paragraph 9 of ownedDoc to style heading1
set font size of font object of text object of paragraph 3 of ownedDoc to 16
set font size of font object of text object of paragraph 5 of ownedDoc to 15
set font size of font object of text object of paragraph 7 of ownedDoc to 15
set font size of font object of text object of paragraph 9 of ownedDoc to 16
my eventLine("headings", "succeeded", "")
on error errorText number errorNumber
my eventLine("headings", "failed", (errorNumber as text) & ": " & errorText)
end try

my eventLine("native_numbering", "attempted", "")
try
set ownList to make new list template at ownedDoc with properties {name:"MSOfficeSpikeOutline", outline numbered:true}
repeat with levelIndex from 1 to 3
set lvl to list level levelIndex of ownList
set number style of lvl to list number style arabic
set start at of lvl to 1
set reset on higher of lvl to (levelIndex - 1)
if levelIndex is 1 then
set number format of lvl to "%1"
else if levelIndex is 2 then
set number format of lvl to "%1.%2"
else
set number format of lvl to "%1.%2.%3"
end if
end repeat
set headingOne to Word style (style heading1) of ownedDoc
set headingTwo to Word style (style heading2) of ownedDoc
set headingThree to Word style (style heading3) of ownedDoc
link to list template headingOne list template ownList list level number 1
link to list template headingTwo list template ownList list level number 2
link to list template headingThree list template ownList list level number 3
my eventLine("native_numbering", "succeeded", "")
on error errorText number errorNumber
my eventLine("native_numbering", "failed", (errorNumber as text) & ": " & errorText)
end try

my eventLine("long_table", "attempted", "")
try
set firstTablePosition to start of content of text object of paragraph 11 of ownedDoc
set finalPosition to (end of content of text object of ownedDoc) - 1
set tableRange to create range ownedDoc start firstTablePosition end finalPosition
set ownedTable to make new table at ownedDoc with properties {text object:tableRange, number of rows:82, number of columns:3}
set rowData to paragraphs of inputText
repeat with rowIndex from 1 to 82
set AppleScript's text item delimiters to tab
set cellData to text items of item (rowIndex + 10) of rowData
set AppleScript's text item delimiters to ""
repeat with columnIndex from 1 to 3
set ownCell to get cell from table ownedTable row rowIndex column columnIndex
set content of text object of ownCell to item columnIndex of cellData
end repeat
end repeat
set allow page breaks of ownedTable to true
set heading format of row 1 of ownedTable to true
set character unit first line indent of paragraph format of text object of ownedTable to 0
my eventLine("table_rows", "observed", (number of rows of ownedTable) as text)
my eventLine("long_table", "succeeded", "")
on error errorText number errorNumber
my eventLine("long_table", "failed", (errorNumber as text) & ": " & errorText)
end try

my eventLine("toc", "attempted", "")
try
set tocPosition to start of content of text object of paragraph 2 of ownedDoc
set tocRange to create range ownedDoc start tocPosition end tocPosition
create new field text range tocRange field type field toc field text "\\o \"1-3\" \\h \\z \\u" preserve formatting true
my eventLine("toc", "succeeded", "")
on error errorText number errorNumber
my eventLine("toc", "failed", (errorNumber as text) & ": " & errorText)
end try

my eventLine("field_refresh", "attempted", "")
try
if (count of fields of ownedDoc) is 0 then error "No fields were created"
repeat with fieldIndex from 1 to (count of fields of ownedDoc)
set ownField to field fieldIndex of ownedDoc
if (update field ownField) is false then error "Word update field returned false"
end repeat
my eventLine("field_refresh", "succeeded", "")
on error errorText number errorNumber
my eventLine("field_refresh", "failed", (errorNumber as text) & ": " & errorText)
end try

my eventLine("repaginate", "attempted", "")
try
repaginate ownedDoc
my eventLine("page_count", "observed", (compute statistics ownedDoc statistic statistic pages) as text)
my eventLine("repaginate", "succeeded", "")
on error errorText number errorNumber
my eventLine("repaginate", "failed", (errorNumber as text) & ": " & errorText)
end try

   my eventLine("save_docx", "attempted", "")
   save as ownedDoc file name docxPath file format format document default add to recent files false
   set ownedDoc to first document whose posix full name is docxPath
   my eventLine("save_docx", "succeeded", posix full name of ownedDoc)
   close ownedDoc saving no
   set ownedDoc to missing value
   my eventLine("reopen", "attempted", "")
   open file name docxPath add to recent files false
   set ownedDoc to first document whose posix full name is docxPath
   if (content of text object of ownedDoc) does not contain "MSOFFICE-SPIKE-END" then error "Reopened document missing end marker"
   my eventLine("reopen", "succeeded", posix full name of ownedDoc)
   set coreOK to true

my eventLine("export_pdf", "attempted", "")
try
save as ownedDoc file name pdfPath file format format PDF add to recent files false
-- PDF SaveAs may change the document identity; reacquire only an exact owned path.
set ownedDoc to missing value
repeat with documentIndex from 1 to (count of documents)
set candidate to document documentIndex
if (posix full name of candidate is docxPath) or (posix full name of candidate is pdfPath) then set ownedDoc to candidate
end repeat
my eventLine("export_pdf", "succeeded", "")
on error errorText number errorNumber
my eventLine("export_pdf", "failed", (errorNumber as text) & ": " & errorText)
end try

  on error errorText number errorNumber
   my eventLine("core", "failed", (errorNumber as text) & ": " & errorText)
  end try
  my eventLine("cleanup", "attempted", "")
  try
   if ownedDoc is not missing value then close ownedDoc saving no
   set ownedDoc to missing value
   my eventLine("cleanup", "succeeded", "Only task-owned document; Word left running")
  on error errorText number errorNumber
   my eventLine("cleanup", "failed", (errorNumber as text) & ": " & errorText)
  end try
  my eventLine("preservation", "attempted", "")
  try
   if snapshotReady is false then error "Initial snapshot incomplete"
   if (count of documents) is not (count beforeDocs) then error "Open-document count changed"
   repeat with prior in beforeDocs
    set matchingDoc to document (item 1 of prior)
    if posix full name of matchingDoc is not item 2 of prior then error "Preexisting document path changed"
    if saved of matchingDoc is not item 3 of prior then error "Preexisting document saved state changed"
    if content of text object of matchingDoc is not item 4 of prior then error "Preexisting document text changed"
    my eventLine("existing_after", "observed", (name of matchingDoc) & "|" & (posix full name of matchingDoc) & "|saved=" & (saved of matchingDoc as text) & "|characters=" & (my textLength(content of text object of matchingDoc)))
   end repeat
   my eventLine("preservation", "succeeded", "Names, exact paths, saved flags and full in-memory text equal")
  on error errorText number errorNumber
   my eventLine("preservation", "failed", (errorNumber as text) & ": " & errorText)
  end try
 end tell
end timeout
if coreOK is false then error "Native core generation/save/reopen failed; see EVENT records"
return "Native probe ended; inspect capability gaps and artifacts separately"
end run
