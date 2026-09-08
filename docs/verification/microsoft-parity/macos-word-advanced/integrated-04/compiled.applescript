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
with timeout of 176 seconds
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
save as ownedDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-x5kazrao/advanced.docx" file format format document default add to recent files false
set ownedDoc to document "advanced.docx"
if posix full name of ownedDoc is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-x5kazrao/advanced.docx" then error "Owned document path mismatch"
   set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.reset"
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.configure_page"
set top margin of page setup of ownedDoc to 72.0
set bottom margin of page setup of ownedDoc to 72.0
set left margin of page setup of ownedDoc to 85.04
set right margin of page setup of ownedDoc to 70.87
set page width of page setup of ownedDoc to 595.28
set page height of page setup of ownedDoc to 841.89
set ownedDoc to document "advanced.docx"
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
set ownStyle to make new Word style at ownedDoc with properties {name local:"TaskColored"}
set base style of ownStyle to style body text
set strike through of font object of ownStyle to true
set paragraph format left indent of paragraph format of ownStyle to 18
set paragraph format right indent of paragraph format of ownStyle to 9
set color of font object of ownStyle to {4369, 21845, 39321}
set underline of font object of ownStyle to underline single
set background pattern color of shading of ownStyle to {59110, 61680, 64250}
set ownBorder to get border (paragraph format of ownStyle) which border border left
set line style of ownBorder to line style single
set line width of ownBorder to line width150 point
set color of ownBorder to {4369, 21845, 39321}
set line spacing rule of paragraph format of ownStyle to line space exactly
set line spacing of paragraph format of ownStyle to 20
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.configure_front_matter"
set nodeStart to (end of content of text object of ownedDoc) - 1
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_f95c46bf5b9f381c21680198", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
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
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.reserve_document_quality_anchor"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "[LONGFORM_NORMALIZATION_ERROR] Semantic normalization failed: 'MathBlock' object has no attribute 'source'." & return)
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_253929dc61580fa5e221a532", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.configure_section"
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_heading"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "Native Word advanced acceptance" & return)
set style of r to style heading1
set pendingHeadingStart to nodeStart
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_heading"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "Native equation" & return)
set style of r to style heading2
set pendingHeadingStart to nodeStart
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_equation"
set nodeStart to (end of content of text object of ownedDoc) - 1
set p to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start p end p
set formulaTable to make new table at ownedDoc with properties {text object:r, number of rows:1, number of columns:3}
set allow break across pages of row 1 of formulaTable to false
set ownBorder to get border formulaTable which border border top
set line style of ownBorder to line style none
set ownBorder to get border formulaTable which border border bottom
set line style of ownBorder to line style none
set ownBorder to get border formulaTable which border border left
set line style of ownBorder to line style none
set ownBorder to get border formulaTable which border border right
set line style of ownBorder to line style none
set ownBorder to get border formulaTable which border border horizontal
set line style of ownBorder to line style none
set ownBorder to get border formulaTable which border border vertical
set line style of ownBorder to line style none
set formulaCell to get cell from table formulaTable row 1 column 2
set content of text object of formulaCell to "x^2+(a)/(b)"
set s to start of content of text object of formulaCell
set e to (end of content of text object of formulaCell) - 1
set r to create range ownedDoc start s end e
set nativeMathCount to count of math objects of ownedDoc
create new equation from range r in document ownedDoc
if (count of math objects of ownedDoc) is not nativeMathCount + 1 then error "Native equation creation failed"
set ownMath to math object (nativeMathCount + 1) of ownedDoc
build up ownMath
set alignment of paragraph format of text object of formulaCell to align paragraph center
set formulaNumberCell to get cell from table formulaTable row 1 column 3
set content of text object of formulaNumberCell to "()"
set numberStart to (start of content of text object of formulaNumberCell) + 1
set r to create range ownedDoc start numberStart end numberStart
create new field text range r field type field sequence field text "WPSC_EQ \\* ARABIC" preserve formatting true
set numberEnd to (end of content of text object of formulaNumberCell) - 2
set r to create range ownedDoc start numberStart end numberEnd
make new bookmark at ownedDoc with properties {name:"wpsc_eq_650de2af4db352e6ca663b09", text object:r}
set alignment of paragraph format of text object of formulaNumberCell to align paragraph right
set first line indent of paragraph format of text object of formulaTable to 0
set keep together of paragraph format of text object of formulaTable to true
set r to my appendText(ownedDoc, return)
set style of r to style body text
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_c78986edff86e7b97b5fd3ca", text object:r}
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "AFTER-EQUATION" & return)
set style of r to style body text
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_heading"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "Landscape merge" & return)
set style of r to style heading2
set pendingHeadingStart to nodeStart
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_semantic_table"
set p to (end of content of text object of ownedDoc) - 1
if pendingHeadingStart is not missing value then set p to pendingHeadingStart
set r to create range ownedDoc start p end p
insert break at r break type section break next page
set nodeStart to (end of content of text object of ownedDoc) - 1
set p to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start p end p
set ownTable to make new table at ownedDoc with properties {text object:r, number of rows:3, number of columns:3}
set allow page breaks of ownTable to true
set heading format of row 1 of ownTable to true
set left padding of ownTable to 0.0
set right padding of ownTable to 0.0
set allow break across pages of row 1 of ownTable to false
set ownCell to get cell from table ownTable row 1 column 1
set content of text object of ownCell to "A"
set alignment of paragraph format of text object of ownCell to align paragraph left
set ownCell to get cell from table ownTable row 1 column 2
set content of text object of ownCell to "B"
set alignment of paragraph format of text object of ownCell to align paragraph left
set ownCell to get cell from table ownTable row 1 column 3
set content of text object of ownCell to "C"
set alignment of paragraph format of text object of ownCell to align paragraph left
set allow break across pages of row 2 of ownTable to false
set ownCell to get cell from table ownTable row 2 column 1
set content of text object of ownCell to "MERGE"
set alignment of paragraph format of text object of ownCell to align paragraph left
set ownCell to get cell from table ownTable row 2 column 2
set content of text object of ownCell to ""
set alignment of paragraph format of text object of ownCell to align paragraph left
set ownCell to get cell from table ownTable row 2 column 3
set content of text object of ownCell to "OUTSIDE"
set alignment of paragraph format of text object of ownCell to align paragraph left
set allow break across pages of row 3 of ownTable to false
set ownCell to get cell from table ownTable row 3 column 1
set content of text object of ownCell to "[REFERENCE_UNRESOLVED 引用目标未解析]"
set alignment of paragraph format of text object of ownCell to align paragraph left
set ownCell to get cell from table ownTable row 3 column 2
set content of text object of ownCell to "middle"
set alignment of paragraph format of text object of ownCell to align paragraph left
set ownCell to get cell from table ownTable row 3 column 3
set content of text object of ownCell to "right"
set alignment of paragraph format of text object of ownCell to align paragraph left
set r to text object of ownTable
set style of r to style body text
set first line indent of paragraph format of r to 0
set character unit first line indent of paragraph format of r to 0
set font size of font object of r to 10
set east asian name of font object of r to "宋体"
set ownBorder to get border ownTable which border border top
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownBorder to get border ownTable which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownBorder to get border ownTable which border border left
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownBorder to get border ownTable which border border right
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownBorder to get border ownTable which border border horizontal
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownBorder to get border ownTable which border border vertical
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownCell to get cell from table ownTable row 1 column 1
set ownBorder to get border ownCell which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownCell to get cell from table ownTable row 1 column 2
set ownBorder to get border ownCell which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownCell to get cell from table ownTable row 1 column 3
set ownBorder to get border ownCell which border border bottom
set line style of ownBorder to line style single
set line width of ownBorder to line width75 point
set ownCell to get cell from table ownTable row 3 column 1
set background pattern color of shading of text object of ownCell to {65535, 60138, 60138}
set mergeStart to get cell from table ownTable row 2 column 1
set mergeEnd to get cell from table ownTable row 2 column 2
merge cell mergeStart with mergeEnd
set r to my appendText(ownedDoc, return)
set nodeEnd to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start nodeStart end nodeEnd
make new bookmark at ownedDoc with properties {name:"wpsc_m5_32caa158ed08f7747eb1f372", text object:r}
set p to (end of content of text object of ownedDoc) - 1
set r to create range ownedDoc start p end p
insert break at r break type section break continuous
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "AFTER-LANDSCAPE" & return)
set style of r to style body text
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_PICTURE_STEP:335"
log "WPSC_OP:writer.add_captioned_figure"
log "WPSC_PICTURE_STEP:336"
set nodeStart to (end of content of text object of ownedDoc) - 1
log "WPSC_PICTURE_STEP:337"
set p to (end of content of text object of ownedDoc) - 1
log "WPSC_PICTURE_STEP:338"
set r to create range ownedDoc start p end p
log "WPSC_PICTURE_STEP:339"
set pictureTable to make new table at ownedDoc with properties {text object:r, number of rows:1, number of columns:2}
log "WPSC_PICTURE_STEP:340"
set ownBorder to get border pictureTable which border border top
log "WPSC_PICTURE_STEP:341"
set line style of ownBorder to line style none
log "WPSC_PICTURE_STEP:342"
set ownBorder to get border pictureTable which border border bottom
log "WPSC_PICTURE_STEP:343"
set line style of ownBorder to line style none
log "WPSC_PICTURE_STEP:344"
set ownBorder to get border pictureTable which border border left
log "WPSC_PICTURE_STEP:345"
set line style of ownBorder to line style none
log "WPSC_PICTURE_STEP:346"
set ownBorder to get border pictureTable which border border right
log "WPSC_PICTURE_STEP:347"
set line style of ownBorder to line style none
log "WPSC_PICTURE_STEP:348"
set ownBorder to get border pictureTable which border border horizontal
log "WPSC_PICTURE_STEP:349"
set line style of ownBorder to line style none
log "WPSC_PICTURE_STEP:350"
set ownBorder to get border pictureTable which border border vertical
log "WPSC_PICTURE_STEP:351"
set line style of ownBorder to line style none
log "WPSC_PICTURE_STEP:352"
set pictureCell to get cell from table pictureTable row 1 column 1
log "WPSC_PICTURE_STEP:353"
set p to (end of content of text object of ownedDoc) - 1
log "WPSC_PICTURE_STEP:354"
set r to create range ownedDoc start p end p
log "WPSC_PICTURE_STEP:355"
set ownPicture to make new inline picture at end of text object of pictureCell with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-x5kazrao/2d211f2c4d5804c0.png", link to file:false, save with document:true}
log "WPSC_PICTURE_STEP:356"
set ownPicture to inline picture (count of inline pictures of ownedDoc) of ownedDoc
log "WPSC_PICTURE_STEP:357"
set width of ownPicture to 180.0
log "WPSC_PICTURE_STEP:358"
set height of ownPicture to 90.0
log "WPSC_PICTURE_STEP:359"
set r to text object of ownPicture
log "WPSC_PICTURE_STEP:360"
set first line indent of paragraph format of r to 0
log "WPSC_PICTURE_STEP:361"
set alignment of paragraph format of r to align paragraph center
log "WPSC_PICTURE_STEP:362"
set keep with next of paragraph format of r to true
log "WPSC_PICTURE_STEP:363"
set pictureCell to get cell from table pictureTable row 1 column 2
log "WPSC_PICTURE_STEP:364"
set p to (end of content of text object of ownedDoc) - 1
log "WPSC_PICTURE_STEP:365"
set r to create range ownedDoc start p end p
log "WPSC_PICTURE_STEP:366"
set ownPicture to make new inline picture at end of text object of pictureCell with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-x5kazrao/9511568f54eeeb65.png", link to file:false, save with document:true}
log "WPSC_PICTURE_STEP:367"
set ownPicture to inline picture (count of inline pictures of ownedDoc) of ownedDoc
log "WPSC_PICTURE_STEP:368"
set width of ownPicture to 180.0
log "WPSC_PICTURE_STEP:369"
set height of ownPicture to 90.0
log "WPSC_PICTURE_STEP:370"
set r to text object of ownPicture
log "WPSC_PICTURE_STEP:371"
set first line indent of paragraph format of r to 0
log "WPSC_PICTURE_STEP:372"
set alignment of paragraph format of r to align paragraph center
log "WPSC_PICTURE_STEP:373"
set keep with next of paragraph format of r to true
log "WPSC_PICTURE_STEP:374"
set r to my appendText(ownedDoc, return)
log "WPSC_PICTURE_STEP:375"
set r to my appendText(ownedDoc, "图 ")
log "WPSC_PICTURE_STEP:376"
set captionStart to (end of content of text object of ownedDoc) - 1
log "WPSC_PICTURE_STEP:377"
set r to create range ownedDoc start captionStart end captionStart
log "WPSC_PICTURE_STEP:378"
create new field text range r field type field sequence field text "WPSC_FIG \\* ARABIC" preserve formatting true
log "WPSC_PICTURE_STEP:379"
set captionEnd to (end of content of text object of ownedDoc) - 1
log "WPSC_PICTURE_STEP:380"
set r to create range ownedDoc start captionStart end captionEnd
log "WPSC_PICTURE_STEP:381"
make new bookmark at ownedDoc with properties {name:"wpsc_fig_72dc80353dd58822363917b0", text object:r}
log "WPSC_PICTURE_STEP:382"
set r to my appendText(ownedDoc, " Two native images" & return)
log "WPSC_PICTURE_STEP:383"
set font size of font object of r to 10
log "WPSC_PICTURE_STEP:384"
set first line indent of paragraph format of r to 0
log "WPSC_PICTURE_STEP:385"
set alignment of paragraph format of r to align paragraph center
log "WPSC_PICTURE_STEP:386"
set keep with next of paragraph format of r to false
log "WPSC_PICTURE_STEP:387"
set nodeEnd to (end of content of text object of ownedDoc) - 1
log "WPSC_PICTURE_STEP:388"
set r to create range ownedDoc start nodeStart end nodeEnd
log "WPSC_PICTURE_STEP:389"
make new bookmark at ownedDoc with properties {name:"wpsc_m5_8126d1f437b53559f31fe7b1", text object:r}
log "WPSC_PICTURE_STEP:390"
set pendingHeadingStart to missing value
log "WPSC_PICTURE_STEP:391"
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "COLORED-STYLE" & return)
set style of r to "TaskColored"
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.add_paragraph"
set nodeStart to (end of content of text object of ownedDoc) - 1
set r to my appendText(ownedDoc, "END-OF-DOCUMENT" & return)
set style of r to style body text
set pendingHeadingStart to missing value
set ownedDoc to document "advanced.docx"
log "WPSC_OP:writer.finalize_fields"
log "WPSC_SECTION_COMMAND:0"
set ownSection to section 1 of ownedDoc
log "WPSC_SECTION_COMMAND:1"
set ownHeader to get header ownSection index header footer primary
log "WPSC_SECTION_COMMAND:2"
set link to previous of ownHeader to false
log "WPSC_SECTION_COMMAND:3"
set content of text object of ownHeader to ""
log "WPSC_SECTION_COMMAND:4"
set ownFooter to get footer ownSection index header footer primary
log "WPSC_SECTION_COMMAND:5"
set link to previous of ownFooter to false
log "WPSC_SECTION_COMMAND:6"
set content of text object of ownFooter to "CONFIDENTIAL "
log "WPSC_SECTION_COMMAND:7"
set number style of page number options of ownFooter to page number style arabic
log "WPSC_SECTION_COMMAND:8"
set restart numbering at section of page number options of ownFooter to true
log "WPSC_SECTION_COMMAND:9"
set starting number of page number options of ownFooter to 1
log "WPSC_SECTION_COMMAND:10"
set alignment of paragraph format of text object of ownFooter to align paragraph center
log "WPSC_SECTION_COMMAND:11"
set content of text object of ownFooter to "CONFIDENTIAL  "
log "WPSC_SECTION_COMMAND:12"
set r to character 14 of text object of ownFooter
log "WPSC_SECTION_COMMAND:13"
create new field text range r field type field page preserve formatting true
log "WPSC_SECTION_COMMAND:14"
set ownSection to section 2 of ownedDoc
log "WPSC_SECTION_COMMAND:15"
set orientation of page setup of ownSection to orient landscape
log "WPSC_SECTION_COMMAND:16"
set mediaFooter to get footer ownSection index header footer primary
log "WPSC_SECTION_COMMAND:17"
set link to previous of mediaFooter to true
log "WPSC_SECTION_COMMAND:18"
set restart numbering at section of page number options of mediaFooter to false
log "WPSC_SECTION_COMMAND:19"
set ownSection to section 3 of ownedDoc
log "WPSC_SECTION_COMMAND:20"
set orientation of page setup of ownSection to orient portrait
log "WPSC_SECTION_COMMAND:21"
set mediaFooter to get footer ownSection index header footer primary
log "WPSC_SECTION_COMMAND:22"
set link to previous of mediaFooter to true
log "WPSC_SECTION_COMMAND:23"
set restart numbering at section of page number options of mediaFooter to false
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
set r to text object of bookmark "wpsc_m5_c78986edff86e7b97b5fd3ca" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "wpsc-eq:1" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_32caa158ed08f7747eb1f372" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "wpsc-tab:1" & tab & sp & tab & pe & tab & s & tab & e
set r to text object of bookmark "wpsc_m5_8126d1f437b53559f31fe7b1" of ownedDoc
set s to start of content of r
set e to end of content of r
set sr to create range ownedDoc start s end s
set ep to e
if ep > s then set ep to ep - 1
set er to create range ownedDoc start ep end ep
set sp to (get range information sr information type active end page number) as integer
set pe to (get range information er information type active end page number) as integer
log "WPSC_NODE" & tab & "wpsc-fig:1" & tab & sp & tab & pe & tab & s & tab & e
save as ownedDoc file name "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-x5kazrao/advanced.docx" file format format document default add to recent files false
set ownedDoc to document "advanced.docx"
  on error errText number errNumber
   log "WPSC_ERROR" & tab & errNumber & tab & errText
   set failureText to "Native Word operation failed (" & errNumber & "): " & errText
  end try
  try
   if ownedDoc is not missing value then
    set closePath to posix full name of ownedDoc
    set closeName to name of ownedDoc
    if closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-x5kazrao/advanced.docx" and closePath is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-native-x5kazrao/advanced.docx" then error "Cleanup identity is outside this operation"
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