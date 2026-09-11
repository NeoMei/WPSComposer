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
set boundDoc to document "document-7d1087ce5166455fbd3591a20be8058c.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-jmrwzigd/document-7d1087ce5166455fbd3591a20be8058c.docx" then error "WPSC_STALE_DOCUMENT"
set nativeRows to {{"stats",compute statistics boundDoc statistic statistic pages}}
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type main text story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:main text/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnotes story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnotes/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnotes story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnotes/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type comments story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:comments/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type text frame story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:text frame/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type even pages header story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:even pages header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type primary header story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:primary header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type even pages footer story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:even pages footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type primary footer story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:primary footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type first page header story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:first page header/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type first page footer story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:first page footer/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnote separator story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote separator/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type 13
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote continuation separator /chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type footnote continuation notice story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:footnote continuation notice/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnote separator story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote separator/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type 16
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote continuation separator /chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
set storyAvailable to false
set ownStory to missing value
try
set ownStory to get story range boundDoc story type endnote continuation notice story
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
set chainIndex to 0
repeat while storyAvailable
set chainIndex to chainIndex + 1
if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set storyOwner to "story:endnote continuation notice/chain:" & chainIndex
repeat with fieldIndex from 1 to count fields of ownStory
set ownField to field fieldIndex of ownStory
set ownKind to ""
set ownType to field type of ownField
if ownType is field page then set ownKind to "PAGE"
if ownType is field num pages then set ownKind to "NUMPAGES"
if ownType is field ref then set ownKind to "REF"
if ownType is field style ref then set ownKind to "STYLEREF"
set codeText to ""
if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text
if ownType is field sequence then
if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"
if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"
if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"
end if
if ownType is field toc then set ownKind to "INDEX"
if ownKind is not "" then
set resultRange to result range of ownField
set nativeStart to start of content of resultRange
set nativeEnd to end of content of resultRange
set firstPage to 0
set lastPage to 0
if ownKind is "INDEX" then
set firstRange to create range boundDoc start nativeStart end nativeStart
set lastPoint to nativeStart
if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1
set lastRange to create range boundDoc start lastPoint end lastPoint
set firstPage to (get range information firstRange information type active end page number) as integer
set lastPage to (get range information lastRange information type active end page number) as integer
end if
set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}
end if
end repeat
set storyAvailable to false
try
set ownStory to next story range of ownStory
on error errorMessage number errorNumber
if errorNumber is not -5941 then error errorMessage number errorNumber
end try
try
set storyAvailable to (ownStory is not missing value)
on error errorMessage number errorNumber
if errorNumber is not -2753 then error errorMessage number errorNumber
end try
end repeat
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
