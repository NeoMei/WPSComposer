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
set boundDoc to document "document-9f9264cfb4884bd6975679df03c1af69.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-povitg9n/document-9f9264cfb4884bd6975679df03c1af69.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {{"stats",compute statistics boundDoc statistic statistic pages}}
if not (exists bookmark "WPSC_F_4bee17b53f19452a8d1d24eb1be500" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set identityRange to text object of bookmark "WPSC_F_4bee17b53f19452a8d1d24eb1be500" of boundDoc
if (content of identityRange as text) is not " TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT " then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"identity","WPSC_F_4bee17b53f19452a8d1d24eb1be500",start of content of identityRange}
if not (exists bookmark "WPSC_F_8e37a4cbf3074205b570605fcfb16f" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set identityRange to text object of bookmark "WPSC_F_8e37a4cbf3074205b570605fcfb16f" of boundDoc
if (content of identityRange as text) is not " TOC \\o \"1-3\" \\h \\z \\* MERGEFORMAT " then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"identity","WPSC_F_8e37a4cbf3074205b570605fcfb16f",start of content of identityRange}
if not (exists bookmark "WPSC_F_3ec92b60a7a54ca68a342128b9c1b9" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set identityRange to text object of bookmark "WPSC_F_3ec92b60a7a54ca68a342128b9c1b9" of boundDoc
if (content of identityRange as text) is not " TOC \\c \"WPSC_FIG\" \\h \\z \\* MERGEFORMAT " then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"identity","WPSC_F_3ec92b60a7a54ca68a342128b9c1b9",start of content of identityRange}
if not (exists bookmark "WPSC_F_623ef084162a4e61bbae25fb8520bc" of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"
set identityRange to text object of bookmark "WPSC_F_623ef084162a4e61bbae25fb8520bc" of boundDoc
if (content of identityRange as text) is not " TOC \\c \"WPSC_TAB\" \\h \\z \\* MERGEFORMAT " then error "WPSC_FIELD_IDENTITY_STALE"
set end of nativeRows to {"identity","WPSC_F_623ef084162a4e61bbae25fb8520bc",start of content of identityRange}
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type main text story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:main text/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnotes story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnotes/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnotes story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnotes/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type comments story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:comments/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type text frame story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:text frame/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type even pages header story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:even pages header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type primary header story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:primary header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type even pages footer story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:even pages footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type primary footer story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:primary footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type first page header story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:first page header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type first page footer story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:first page footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnote separator story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote separator/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type 13
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote continuation separator /chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnote continuation notice story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote continuation notice/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnote separator story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote separator/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type 16
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote continuation separator /chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnote continuation notice story
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote continuation notice/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
if field type of ownField is field page then set ownKind to "PAGE"
if field type of ownField is field num pages then set ownKind to "NUMPAGES"
if field type of ownField is field ref then set ownKind to "REF"
if field type of ownField is field style ref then set ownKind to "STYLEREF"
set codeText to content of field code of ownField as text
if field type of ownField is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if field type of ownField is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstRange to set range resultRange start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to set range resultRange start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -5941 and errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
