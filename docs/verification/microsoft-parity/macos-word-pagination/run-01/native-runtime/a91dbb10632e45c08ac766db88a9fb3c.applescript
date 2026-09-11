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
set boundDoc to document "document-dd8978c7273641f49540cfdcb2d4b494.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-oocs124i/document-dd8978c7273641f49540cfdcb2d4b494.docx" then error "WPSC_STALE_DOCUMENT"
repaginate boundDoc
set paginationSetup to page setup of boundDoc
set paginationWidth to 595.28
try
set paginationWidth to page width of paginationSetup
on error paginationError number paginationNumber
if paginationNumber is not -1728 then error paginationError number paginationNumber
end try
set paginationHeight to 841.89
try
set paginationHeight to page height of paginationSetup
on error paginationError number paginationNumber
if paginationNumber is not -1728 then error paginationError number paginationNumber
end try
set paginationLeft to 72.0
try
set paginationLeft to left margin of paginationSetup
on error paginationError number paginationNumber
if paginationNumber is not -1728 then error paginationError number paginationNumber
end try
set paginationRight to 72.0
try
set paginationRight to right margin of paginationSetup
on error paginationError number paginationNumber
if paginationNumber is not -1728 then error paginationError number paginationNumber
end try
set paginationTop to 72.0
try
set paginationTop to top margin of paginationSetup
on error paginationError number paginationNumber
if paginationNumber is not -1728 then error paginationError number paginationNumber
end try
set paginationBottom to 72.0
try
set paginationBottom to bottom margin of paginationSetup
on error paginationError number paginationNumber
if paginationNumber is not -1728 then error paginationError number paginationNumber
end try
set nativeRows to {{"setup",paginationWidth,paginationHeight,paginationLeft,paginationRight,paginationTop,paginationBottom}}
set paginationContentEnd to (end of content of text object of boundDoc) - 1
if paginationContentEnd < 0 then set paginationContentEnd to 0
set paginationFirst to 0
if paginationFirst > paginationContentEnd then set paginationFirst to paginationContentEnd
set paginationLast to 92
if paginationLast > paginationContentEnd then set paginationLast to paginationContentEnd
set paginationFirstRange to create range boundDoc start paginationFirst end paginationFirst
set paginationLastRange to create range boundDoc start paginationLast end paginationLast
set paginationFirstPage to get range information paginationFirstRange information type active end page number
set paginationLastPage to get range information paginationLastRange information type active end page number
set paginationFirstX to get range information paginationFirstRange information type horizontal position relative to page
set paginationFirstY to get range information paginationFirstRange information type vertical position relative to page
set paginationLastY to get range information paginationLastRange information type vertical position relative to page
set end of nativeRows to {"range",0,paginationFirstPage,paginationLastPage,paginationFirstX,paginationFirstY,paginationLastY}
set paginationFirst to 0
if paginationFirst > paginationContentEnd then set paginationFirst to paginationContentEnd
set paginationLast to 1
if paginationLast > paginationContentEnd then set paginationLast to paginationContentEnd
set paginationFirstRange to create range boundDoc start paginationFirst end paginationFirst
set paginationLastRange to create range boundDoc start paginationLast end paginationLast
set paginationFirstPage to get range information paginationFirstRange information type active end page number
set paginationLastPage to get range information paginationLastRange information type active end page number
set paginationFirstX to get range information paginationFirstRange information type horizontal position relative to page
set paginationFirstY to get range information paginationFirstRange information type vertical position relative to page
set paginationLastY to get range information paginationLastRange information type vertical position relative to page
set end of nativeRows to {"range",1,paginationFirstPage,paginationLastPage,paginationFirstX,paginationFirstY,paginationLastY}
set paginationFirst to 92
if paginationFirst > paginationContentEnd then set paginationFirst to paginationContentEnd
set paginationLast to 92
if paginationLast > paginationContentEnd then set paginationLast to paginationContentEnd
set paginationFirstRange to create range boundDoc start paginationFirst end paginationFirst
set paginationLastRange to create range boundDoc start paginationLast end paginationLast
set paginationFirstPage to get range information paginationFirstRange information type active end page number
set paginationLastPage to get range information paginationLastRange information type active end page number
set paginationFirstX to get range information paginationFirstRange information type horizontal position relative to page
set paginationFirstY to get range information paginationFirstRange information type vertical position relative to page
set paginationLastY to get range information paginationLastRange information type vertical position relative to page
set end of nativeRows to {"range",2,paginationFirstPage,paginationLastPage,paginationFirstX,paginationFirstY,paginationLastY}
set paginationFirst to 98
if paginationFirst > paginationContentEnd then set paginationFirst to paginationContentEnd
set paginationLast to 101
if paginationLast > paginationContentEnd then set paginationLast to paginationContentEnd
set paginationFirstRange to create range boundDoc start paginationFirst end paginationFirst
set paginationLastRange to create range boundDoc start paginationLast end paginationLast
set paginationFirstPage to get range information paginationFirstRange information type active end page number
set paginationLastPage to get range information paginationLastRange information type active end page number
set paginationFirstX to get range information paginationFirstRange information type horizontal position relative to page
set paginationFirstY to get range information paginationFirstRange information type vertical position relative to page
set paginationLastY to get range information paginationLastRange information type vertical position relative to page
set end of nativeRows to {"range",3,paginationFirstPage,paginationLastPage,paginationFirstX,paginationFirstY,paginationLastY}
set paginationFirst to 0
if paginationFirst > paginationContentEnd then set paginationFirst to paginationContentEnd
set paginationLast to 0
if paginationLast > paginationContentEnd then set paginationLast to paginationContentEnd
set paginationFirstRange to create range boundDoc start paginationFirst end paginationFirst
set paginationLastRange to create range boundDoc start paginationLast end paginationLast
set paginationFirstPage to get range information paginationFirstRange information type active end page number
set paginationLastPage to get range information paginationLastRange information type active end page number
set paginationFirstX to get range information paginationFirstRange information type horizontal position relative to page
set paginationFirstY to get range information paginationFirstRange information type vertical position relative to page
set paginationLastY to get range information paginationLastRange information type vertical position relative to page
set end of nativeRows to {"range",4,paginationFirstPage,paginationLastPage,paginationFirstX,paginationFirstY,paginationLastY}
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
