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
set boundDoc to document "document-b2df23ac535642cea95b0fa58ec3eec4.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5d4f05bp/document-b2df23ac535642cea95b0fa58ec3eec4.docx" then error "WPSC_STALE_DOCUMENT"
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
if not ((current application's NSArray's arrayWithArray:nativeRows)'s isEqualToArray:{{"body","HEAD" & return & "left REPLACE-长😀 right" & return & "TAIL" & return & "" & return & ""},{"selection",10,21,"REPLACE-长😀"},{"bookmark","figure_probe_target",10,21,"REPLACE-长😀"},{"bookmark","figure_probe_suffix",21,27," right"},{"format",17,3,5,11,7,21,true,14,true,false,false},{"counts",0,0,0,0,2}}) then error "WPSC_FIGURE_PROBE_PREIMAGE_CHANGED"
set nativeRows to {}
set probeStart to 10
set probeOriginalText to "REPLACE-长😀"
set probeTarget to text object of bookmark "figure_probe_target" of boundDoc
set probeOriginalNoProofing to no proofing of probeTarget
if probeOriginalNoProofing is not false then error "WPSC_FIGURE_PROBE_PROOFING_PREIMAGE"
set probeMutation to create range boundDoc start probeStart end (end of content of probeTarget)
set content of probeMutation to ""
set probeInsertion to create range boundDoc start probeStart end probeStart
set probeBeforePictures to count inline pictures of boundDoc
set probePicture to make new inline picture at probeInsertion with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5d4f05bp/image-b8ebaf076b1c4fcab873fb736d9eb8a4.png",link to file:false,save with document:true}
set probePictureIndex to count inline pictures of boundDoc
set probePicture to inline picture probePictureIndex of boundDoc
set lock aspect ratio of probePicture to true
set width of probePicture to 80
set alternative text of probePicture to "figure:probe/second-failure"
set probePictureRange to text object of probePicture
set alignment of paragraph format of probePictureRange to align paragraph center
set keep together of paragraph format of probePictureRange to true
set keep with next of paragraph format of probePictureRange to true
set probePictureStart to start of content of probePictureRange
set probePictureEnd to end of content of probePictureRange
set probeBreak to create range boundDoc start probePictureEnd end probePictureEnd
set content of probeBreak to return
set probeMutationEnd to probePictureEnd + 1
set selection start of selection of boundWindow to probeMutationEnd
set selection end of selection of boundWindow to probeMutationEnd
set probeSecondFailed to false
try
set probeSecondInsertion to create range boundDoc start probeMutationEnd end probeMutationEnd
make new inline picture at probeSecondInsertion with properties {file name:"/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-5d4f05bp/definitely-missing-figure-probe.png",link to file:false,save with document:true}
on error
set probeSecondFailed to true
end try
if probeSecondFailed is false then error "WPSC_FIGURE_PROBE_EXPECTED_SECOND_FAILURE"
if (count inline pictures of boundDoc) is not probeBeforePictures + 1 then error "WPSC_FIGURE_PROBE_SECOND_FAILURE_SIDE_EFFECT"
set probeSuffixBookmark to bookmark "figure_probe_suffix" of boundDoc
set probeRollbackEnd to probeMutationEnd
if probePictureStart is not 10 or probePictureEnd is not 11 or probeRollbackEnd is not 12 then error "WPSC_FIGURE_PROBE_BOUND_CHANGED"
if content of text object of boundDoc is not "HEAD" & return & "left /" & return & " right" & return & "TAIL" & return & "" & return & "" then error "WPSC_FIGURE_PROBE_BODY_CHANGED"
if start of bookmark of probeSuffixBookmark is not 10 or end of bookmark of probeSuffixBookmark is not 18 then error "WPSC_FIGURE_PROBE_SUFFIX_BOUNDS"
if content of text object of probeSuffixBookmark is not "/" & return & " right" then error "WPSC_FIGURE_PROBE_SUFFIX_CHANGED"
if (count inline pictures of boundDoc) is not 1 then error "WPSC_FIGURE_PROBE_PICTURE_COUNT"
if alternative text of inline picture 1 of boundDoc is not "figure:probe/second-failure" then error "WPSC_FIGURE_PROBE_PICTURE_IDENTITY"
set nativeRows to {{"inserted","second-failure",probePictureStart,probePictureEnd,probePictureIndex,probeBeforePictures,alternative text of probePicture as text,width of probePicture,height of probePicture,(alignment of paragraph format of probePictureRange is align paragraph center),keep together of paragraph format of probePictureRange,keep with next of paragraph format of probePictureRange,probeMutationEnd,probeRollbackEnd,true}}
set probeRollback to create range boundDoc start probeStart end probeRollbackEnd
set content of probeRollback to probeOriginalText
if exists bookmark "figure_probe_target" of boundDoc then delete bookmark "figure_probe_target" of boundDoc
set probeRestored to create range boundDoc start 10 end 21
make new bookmark at boundDoc with properties {name:"figure_probe_target",text object:probeRestored}
if exists bookmark "figure_probe_suffix" of boundDoc then delete bookmark "figure_probe_suffix" of boundDoc
set probeRestoredSuffix to create range boundDoc start 21 end 27
make new bookmark at boundDoc with properties {name:"figure_probe_suffix",text object:probeRestoredSuffix}
set first line indent of paragraph format of probeRestored to 17
set paragraph format left indent of paragraph format of probeRestored to 3
set paragraph format right indent of paragraph format of probeRestored to 5
set space before of paragraph format of probeRestored to 11
set space after of paragraph format of probeRestored to 7
set line spacing rule of paragraph format of probeRestored to line space exactly
set line spacing of paragraph format of probeRestored to 21
set italic of font object of probeRestored to true
set font size of font object of probeRestored to 14
set no proofing of probeRestored to probeOriginalNoProofing
if no proofing of probeRestored is not probeOriginalNoProofing then error "WPSC_FIGURE_PROBE_PROOFING_RESTORE"
set alignment of paragraph format of probeRestored to align paragraph right
set keep together of paragraph format of probeRestored to false
set keep with next of paragraph format of probeRestored to false
set selection start of selection of boundWindow to 10
set selection end of selection of boundWindow to 21
set probeSelection to text object of selection of boundWindow
set probeTargetBookmark to bookmark "figure_probe_target" of boundDoc
set probeSuffixBookmark to bookmark "figure_probe_suffix" of boundDoc
set probeFormatRange to text object of probeTargetBookmark
set end of nativeRows to {"body",content of text object of boundDoc as text}
set end of nativeRows to {"selection",start of content of probeSelection,end of content of probeSelection,content of probeSelection as text}
set end of nativeRows to {"bookmark",name of probeTargetBookmark,start of bookmark of probeTargetBookmark,end of bookmark of probeTargetBookmark,content of text object of probeTargetBookmark as text}
set end of nativeRows to {"bookmark",name of probeSuffixBookmark,start of bookmark of probeSuffixBookmark,end of bookmark of probeSuffixBookmark,content of text object of probeSuffixBookmark as text}
set end of nativeRows to {"format",first line indent of paragraph format of probeFormatRange,paragraph format left indent of paragraph format of probeFormatRange,paragraph format right indent of paragraph format of probeFormatRange,space before of paragraph format of probeFormatRange,space after of paragraph format of probeFormatRange,line spacing of paragraph format of probeFormatRange,italic of font object of probeFormatRange,font size of font object of probeFormatRange,(alignment of paragraph format of probeFormatRange is align paragraph right),keep together of paragraph format of probeFormatRange,keep with next of paragraph format of probeFormatRange}
set end of nativeRows to {"counts",count inline pictures of boundDoc,count shapes of boundDoc,count tables of boundDoc,count fields of boundDoc,count bookmarks of boundDoc}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
