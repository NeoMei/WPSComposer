on appendText(ownedDoc, valueText)
 tell application "Microsoft Word"
  set p to (end of content of text object of ownedDoc) - 1
  set r to create range ownedDoc start p end p
  set content of r to valueText
  set e to (end of content of text object of ownedDoc) - 1
  return create range ownedDoc start p end e
 end tell
end appendText
set ownedDoc to missing value
set beforeDocs to {}
set failureText to ""
with timeout of 237 seconds
 tell application "Microsoft Word"
  repeat with documentIndex from 1 to (count of documents)
   set d to document documentIndex
   set end of beforeDocs to {name of d, posix full name of d, saved of d, content of text object of d}
  end repeat
  try
   set ownedDoc to make new document
set ownedName to name of ownedDoc
repeat with prior in beforeDocs
 if item 1 of prior is ownedName then
  set ownedDoc to missing value
  error "New document identity collision"
 end if
end repeat
set ownedDoc to document ownedName
save as ownedDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-pjrn6xas/document-b2db1d7d39394ece9b0455dc9397c635.docx" file format format document default add to recent files false
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
if posix full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-pjrn6xas/document-b2db1d7d39394ece9b0455dc9397c635.docx" then error "Owned document path mismatch"
   set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.reset"
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.configure_page"
set top margin of page setup of ownedDoc to 72.0
set bottom margin of page setup of ownedDoc to 72.0
set left margin of page setup of ownedDoc to 85.04
set right margin of page setup of ownedDoc to 70.87
set page width of page setup of ownedDoc to 595.28
set page height of page setup of ownedDoc to 841.89
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.ensure_styles"
set ownStyle to Word style (style body text) of ownedDoc
set east asian name of font object of ownStyle to "仿宋"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 12
set first line indent of paragraph format of ownStyle to 24
set alignment of paragraph format of ownStyle to align paragraph justify
set line spacing rule of paragraph format of ownStyle to line space1 pt5
set ownStyle to Word style (style title) of ownedDoc
set east asian name of font object of ownStyle to "黑体"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 22
set bold of font object of ownStyle to true
set alignment of paragraph format of ownStyle to align paragraph center
set outline level of paragraph format of ownStyle to outline level body text
set ownStyle to Word style (style heading1) of ownedDoc
set east asian name of font object of ownStyle to "黑体"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 16
set bold of font object of ownStyle to true
set space before of paragraph format of ownStyle to 16
set space after of paragraph format of ownStyle to 5
set keep with next of paragraph format of ownStyle to true
set alignment of paragraph format of ownStyle to align paragraph center
set outline level of paragraph format of ownStyle to outline level1
set ownStyle to Word style (style heading2) of ownedDoc
set east asian name of font object of ownStyle to "黑体"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 15
set bold of font object of ownStyle to true
set space before of paragraph format of ownStyle to 14
set space after of paragraph format of ownStyle to 5
set keep with next of paragraph format of ownStyle to true
set alignment of paragraph format of ownStyle to align paragraph left
set outline level of paragraph format of ownStyle to outline level2
set ownStyle to Word style (style heading3) of ownedDoc
set east asian name of font object of ownStyle to "黑体"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 15
set bold of font object of ownStyle to true
set space before of paragraph format of ownStyle to 12
set space after of paragraph format of ownStyle to 5
set keep with next of paragraph format of ownStyle to true
set alignment of paragraph format of ownStyle to align paragraph left
set outline level of paragraph format of ownStyle to outline level3
set ownStyle to Word style (style heading4) of ownedDoc
set east asian name of font object of ownStyle to "黑体"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 14
set bold of font object of ownStyle to true
set space before of paragraph format of ownStyle to 10
set space after of paragraph format of ownStyle to 5
set keep with next of paragraph format of ownStyle to true
set alignment of paragraph format of ownStyle to align paragraph left
set outline level of paragraph format of ownStyle to outline level4
set ownStyle to Word style (style heading5) of ownedDoc
set east asian name of font object of ownStyle to "黑体"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 14
set bold of font object of ownStyle to true
set space before of paragraph format of ownStyle to 8
set space after of paragraph format of ownStyle to 5
set keep with next of paragraph format of ownStyle to true
set alignment of paragraph format of ownStyle to align paragraph left
set outline level of paragraph format of ownStyle to outline level5
set ownStyle to Word style (style heading6) of ownedDoc
set east asian name of font object of ownStyle to "黑体"
set name of font object of ownStyle to "Times New Roman"
set font size of font object of ownStyle to 12
set bold of font object of ownStyle to true
set space before of paragraph format of ownStyle to 6
set space after of paragraph format of ownStyle to 5
set keep with next of paragraph format of ownStyle to true
set alignment of paragraph format of ownStyle to align paragraph left
set outline level of paragraph format of ownStyle to outline level6
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.configure_front_matter"
set nodeStart to (end of content of text object of ownedDoc) - 1
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_f95c46bf5b9f381c21680198", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.configure_toc_styles"
set ownStyle to Word style (style toc1) of ownedDoc
set font size of font object of ownStyle to 10.5
set space before of paragraph format of ownStyle to 0.0
set space after of paragraph format of ownStyle to 0.0
set ownStyle to Word style (style toc2) of ownedDoc
set font size of font object of ownStyle to 10.0
set space before of paragraph format of ownStyle to 0.0
set space after of paragraph format of ownStyle to 0.0
set ownStyle to Word style (style toc3) of ownedDoc
set font size of font object of ownStyle to 10.0
set space before of paragraph format of ownStyle to 0.0
set space after of paragraph format of ownStyle to 0.0
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.reserve_document_quality_anchor"
set nodeStart to (end of content of text object of ownedDoc) - 1
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_253929dc61580fa5e221a532", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.configure_section"
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "LIST ARGUMENT NATIVE ACCEPTANCE" & return)
set style of r to style title
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_20de0b04f12ea90041e4475e", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "BASE BODY" & return)
set style of r to style body text
set alignment of paragraph format of r to align paragraph right
set spanRange to create range ownedDoc start (nodeStart + 0) end (nodeStart + 9)
set italic of font object of spanRange to true
set strike through of font object of spanRange to true
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_3a3fba836afcbbd168752b3c", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_list"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "•" & tab & "DEFAULT ITEM" & return)
set style of r to style list paragraph
reset font object of r
reset paragraph format of r
set east asian name of font object of r to "仿宋"
set name of font object of r to "Times New Roman"
set font size of font object of r to 12
set bold of font object of r to false
set first line indent of paragraph format of r to -24
set space before of paragraph format of r to 0
set space after of paragraph format of r to 3
set paragraph format left indent of paragraph format of r to 24
set color of font object of r to {0, 0, 0}
set line spacing rule of paragraph format of r to line space1 pt5
make new tab stop at paragraph 1 of r with properties {tab stop position:24}
set trailingPoint to (end of content of text object of ownedDoc) - 1
set trailingRange to create range ownedDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_6683359dd3f3e6864dfe4197", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "BODY AFTER DEFAULT" & return)
set style of r to style body text
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_a1ce104a326e75e78f56ed75", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_list"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "→" & tab & "CUSTOM ITEM" & return)
set style of r to style list paragraph
reset font object of r
reset paragraph format of r
set east asian name of font object of r to "仿宋"
set name of font object of r to "Times New Roman"
set font size of font object of r to 12
set bold of font object of r to false
set first line indent of paragraph format of r to -30
set space before of paragraph format of r to 0
set space after of paragraph format of r to 3
set paragraph format left indent of paragraph format of r to 30
set color of font object of r to {0, 0, 0}
set line spacing rule of paragraph format of r to line space1 pt5
make new tab stop at paragraph 1 of r with properties {tab stop position:30}
set trailingPoint to (end of content of text object of ownedDoc) - 1
set trailingRange to create range ownedDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_d6918a3ab3f05a0435d18485", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_list"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "" & tab & "EMPTY GLYPH ITEM" & return)
set style of r to style list paragraph
reset font object of r
reset paragraph format of r
set east asian name of font object of r to "仿宋"
set name of font object of r to "Times New Roman"
set font size of font object of r to 12
set bold of font object of r to false
set first line indent of paragraph format of r to -27.5
set space before of paragraph format of r to 0
set space after of paragraph format of r to 3
set paragraph format left indent of paragraph format of r to 27.5
set color of font object of r to {0, 0, 0}
set line spacing rule of paragraph format of r to line space1 pt5
make new tab stop at paragraph 1 of r with properties {tab stop position:27.5}
set trailingPoint to (end of content of text object of ownedDoc) - 1
set trailingRange to create range ownedDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_3a8f5c56395b6ceb22e540d8", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_list"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "1." & tab & "ORDERED ITEM" & return)
set style of r to style list paragraph
reset font object of r
reset paragraph format of r
set east asian name of font object of r to "仿宋"
set name of font object of r to "Times New Roman"
set font size of font object of r to 12
set bold of font object of r to false
set first line indent of paragraph format of r to -33
set space before of paragraph format of r to 0
set space after of paragraph format of r to 3
set paragraph format left indent of paragraph format of r to 33
set color of font object of r to {0, 0, 0}
set line spacing rule of paragraph format of r to line space1 pt5
make new tab stop at paragraph 1 of r with properties {tab stop position:33}
set trailingPoint to (end of content of text object of ownedDoc) - 1
set trailingRange to create range ownedDoc start trailingPoint end trailingPoint
set style of trailingRange to style normal
reset font object of trailingRange
reset paragraph format of trailingRange
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_1101744dad8c5b85cbee4ab3", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "BODY AFTER ORDERED" & return)
set style of r to style body text
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_5d990730a7120a2cc0b8ba6e", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
log "WPSC_OP:writer.finalize_fields"
log "WPSC_SECTION_COMMAND:0"
set ownSection to section 1 of ownedDoc
log "WPSC_SECTION_COMMAND:1"
set ownHeader to get header ownSection index header footer primary
log "WPSC_SECTION_COMMAND:2"
set link to previous of ownHeader to false
log "WPSC_SECTION_COMMAND:3"
set content of text object of ownHeader to "LIST ARGUMENT NATIVE ACCEPTANCE"
log "WPSC_SECTION_COMMAND:4"
set ownFooter to get footer ownSection index header footer primary
log "WPSC_SECTION_COMMAND:5"
set link to previous of ownFooter to false
log "WPSC_SECTION_COMMAND:6"
set content of text object of ownFooter to ""
log "WPSC_SECTION_COMMAND:7"
set number style of page number options of ownFooter to page number style arabic
log "WPSC_SECTION_COMMAND:8"
set restart numbering at section of page number options of ownFooter to true
log "WPSC_SECTION_COMMAND:9"
set starting number of page number options of ownFooter to 1
log "WPSC_SECTION_COMMAND:10"
set alignment of paragraph format of text object of ownFooter to align paragraph center
log "WPSC_SECTION_COMMAND:11"
set content of text object of ownFooter to " "
log "WPSC_SECTION_COMMAND:12"
set r to character 1 of text object of ownFooter
log "WPSC_SECTION_COMMAND:13"
create new field text range r field type field page preserve formatting true
set priorFields to missing value
set fieldsStable to false
repeat with refreshRound from 1 to 3
 repaginate ownedDoc
 repeat with tocIndex from 1 to (count of tables of contents of ownedDoc)
  update (table of contents tocIndex of ownedDoc)
 end repeat
 repeat with figuresIndex from 1 to (count of tables of figures of ownedDoc)
  update (table of figures figuresIndex of ownedDoc)
 end repeat
 repeat with nativeIndex from 1 to (count of indexes of ownedDoc)
  update (index nativeIndex of ownedDoc)
 end repeat
 repeat with fieldIndex from 1 to (count of fields of ownedDoc)
  if (update field (field fieldIndex of ownedDoc)) is false then error "Native Word field update failed"
 end repeat
 repaginate ownedDoc
 set fieldState to {compute statistics ownedDoc statistic statistic pages}
 repeat with fieldIndex from 1 to (count of fields of ownedDoc)
  set ownField to field fieldIndex of ownedDoc
  set end of fieldState to content of result range of ownField
 end repeat
 if fieldState is priorFields then
  set fieldsStable to true
  exit repeat
 end if
 set priorFields to fieldState
end repeat
if fieldsStable is false then error "Native Word fields did not converge"

repaginate ownedDoc
set r to text object of bookmark "wpsc_m5_f95c46bf5b9f381c21680198" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "doc:front-matter" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_253929dc61580fa5e221a532" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "doc:quality" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_20de0b04f12ea90041e4475e" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "doc:title-display" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_3a3fba836afcbbd168752b3c" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "__wpsc_para:1:1" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_6683359dd3f3e6864dfe4197" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "list-case:1" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_a1ce104a326e75e78f56ed75" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "body-after-default" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_d6918a3ab3f05a0435d18485" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "list-case:2" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_3a8f5c56395b6ceb22e540d8" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "list-case:3" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_1101744dad8c5b85cbee4ab3" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "list-case:4" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_5d990730a7120a2cc0b8ba6e" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "body-after-ordered" & tab & sp & tab & pe & tab & s & tab & e
save as ownedDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-pjrn6xas/document-b2db1d7d39394ece9b0455dc9397c635.docx" file format format document default add to recent files false
set ownedDoc to document "document-b2db1d7d39394ece9b0455dc9397c635.docx"
  on error errText number errNumber
   log "WPSC_ERROR" & tab & errNumber & tab & errText
   set failureText to "Native Word operation failed (" & errNumber & "): " & errText
  end try
  try
   if ownedDoc is not missing value then
    set closePath to posix full name of ownedDoc
    set closeName to name of ownedDoc
    if closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-pjrn6xas/document-b2db1d7d39394ece9b0455dc9397c635.docx" and closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-pjrn6xas/document-b2db1d7d39394ece9b0455dc9397c635.docx" then error "Cleanup identity is outside this operation"
    close document closeName saving no
   end if
   set ownedDoc to missing value
  on error
   error "Native Word owned-document cleanup failed; staging quarantined"
  end try
  if (count of documents) is not (count of beforeDocs) then error "Native Word document count changed; staging quarantined"
  repeat with prior in beforeDocs
   set d to document (item 1 of prior)
   if posix full name of d is not item 2 of prior then error "Preexisting Word path changed"
   if saved of d is not item 3 of prior then error "Preexisting Word saved state changed"
   if content of text object of d is not item 4 of prior then error "Preexisting Word text changed"
  end repeat
  log "WPSC_CLEAN"
  if failureText is not "" then error failureText
 end tell
end timeout
return "WPSC_OK" & tab & (count of beforeDocs)
