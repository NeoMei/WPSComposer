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
set boundDoc to document "document-f787378a23fb491ca54f59a3436dbcdc.docx"
set boundWindow to active window of boundDoc
if (posix full name of boundDoc as text) is not "/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s70q3sly/document-f787378a23fb491ca54f59a3436dbcdc.docx" then error "WPSC_STALE_DOCUMENT"
set qualitySelection to selection of boundWindow
if not ((current application's NSString's stringWithString:(posix full name of boundDoc as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s70q3sly/document-f787378a23fb491ca54f59a3436dbcdc.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of boundWindow as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s70q3sly/document-f787378a23fb491ca54f59a3436dbcdc.docx")) then error "WPSC_STALE_DOCUMENT"
if not ((current application's NSString's stringWithString:(posix full name of document of qualitySelection as text))'s isEqualToString:("/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-s70q3sly/document-f787378a23fb491ca54f59a3436dbcdc.docx")) then error "WPSC_STALE_DOCUMENT"
set qualityPoint to 111
set qualityDelta to 0
set qualityState to {}
set qualityLayout to {}
set probeRows to {}
set probePhase to "start"
try
set probePhase to "story-availability-and-chain:0: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "story-availability-and-chain:1: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:2: try"
log probePhase
try
set probePhase to "story-availability-and-chain:3: set probeStoryRange to get story range boundDoc story type primary header story"
log probePhase
set probeStoryRange to get story range boundDoc story type primary header story
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:5: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:7: try"
log probePhase
try
set probePhase to "story-availability-and-chain:8: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:10: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:12: set end of qualityLayout to {\"page-story-available\",\"primary header\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","primary header",probeStoryAvailable as boolean}
set probePhase to "story-availability-and-chain:13: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "story-availability-and-chain:14: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "story-availability-and-chain:15: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "story-availability-and-chain:16: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "story-availability-and-chain:17: set end of qualityLayout to {\"page-story\",\"primary header\",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of conte"
log probePhase
set end of qualityLayout to {"page-story","primary header",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of content of probeStoryRange as integer,end of content of probeStoryRange as integer,content of probeStoryRange as text}
set probePhase to "story-availability-and-chain:18: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "story-availability-and-chain:19: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "story-availability-and-chain:20: set end of qualityLayout to {\"page-field\",\"primary header\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,c"
log probePhase
set end of qualityLayout to {"page-field","primary header",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "story-availability-and-chain:22: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:23: try"
log probePhase
try
set probePhase to "story-availability-and-chain:24: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:26: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:28: try"
log probePhase
try
set probePhase to "story-availability-and-chain:29: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:31: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "story-availability-and-chain:34: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "story-availability-and-chain:35: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:36: try"
log probePhase
try
set probePhase to "story-availability-and-chain:37: set probeStoryRange to get story range boundDoc story type primary footer story"
log probePhase
set probeStoryRange to get story range boundDoc story type primary footer story
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:39: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:41: try"
log probePhase
try
set probePhase to "story-availability-and-chain:42: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:44: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:46: set end of qualityLayout to {\"page-story-available\",\"primary footer\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","primary footer",probeStoryAvailable as boolean}
set probePhase to "story-availability-and-chain:47: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "story-availability-and-chain:48: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "story-availability-and-chain:49: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "story-availability-and-chain:50: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "story-availability-and-chain:51: set end of qualityLayout to {\"page-story\",\"primary footer\",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of conte"
log probePhase
set end of qualityLayout to {"page-story","primary footer",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of content of probeStoryRange as integer,end of content of probeStoryRange as integer,content of probeStoryRange as text}
set probePhase to "story-availability-and-chain:52: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "story-availability-and-chain:53: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "story-availability-and-chain:54: set end of qualityLayout to {\"page-field\",\"primary footer\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,c"
log probePhase
set end of qualityLayout to {"page-field","primary footer",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "story-availability-and-chain:56: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:57: try"
log probePhase
try
set probePhase to "story-availability-and-chain:58: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:60: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:62: try"
log probePhase
try
set probePhase to "story-availability-and-chain:63: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:65: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "story-availability-and-chain:68: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "story-availability-and-chain:69: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:70: try"
log probePhase
try
set probePhase to "story-availability-and-chain:71: set probeStoryRange to get story range boundDoc story type first page header story"
log probePhase
set probeStoryRange to get story range boundDoc story type first page header story
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:73: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:75: try"
log probePhase
try
set probePhase to "story-availability-and-chain:76: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:78: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:80: set end of qualityLayout to {\"page-story-available\",\"first page header\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","first page header",probeStoryAvailable as boolean}
set probePhase to "story-availability-and-chain:81: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "story-availability-and-chain:82: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "story-availability-and-chain:83: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "story-availability-and-chain:84: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "story-availability-and-chain:85: set end of qualityLayout to {\"page-story\",\"first page header\",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of co"
log probePhase
set end of qualityLayout to {"page-story","first page header",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of content of probeStoryRange as integer,end of content of probeStoryRange as integer,content of probeStoryRange as text}
set probePhase to "story-availability-and-chain:86: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "story-availability-and-chain:87: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "story-availability-and-chain:88: set end of qualityLayout to {\"page-field\",\"first page header\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","first page header",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "story-availability-and-chain:90: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:91: try"
log probePhase
try
set probePhase to "story-availability-and-chain:92: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:94: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:96: try"
log probePhase
try
set probePhase to "story-availability-and-chain:97: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:99: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "story-availability-and-chain:102: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "story-availability-and-chain:103: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:104: try"
log probePhase
try
set probePhase to "story-availability-and-chain:105: set probeStoryRange to get story range boundDoc story type first page footer story"
log probePhase
set probeStoryRange to get story range boundDoc story type first page footer story
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:107: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:109: try"
log probePhase
try
set probePhase to "story-availability-and-chain:110: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:112: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:114: set end of qualityLayout to {\"page-story-available\",\"first page footer\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","first page footer",probeStoryAvailable as boolean}
set probePhase to "story-availability-and-chain:115: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "story-availability-and-chain:116: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "story-availability-and-chain:117: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "story-availability-and-chain:118: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "story-availability-and-chain:119: set end of qualityLayout to {\"page-story\",\"first page footer\",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of co"
log probePhase
set end of qualityLayout to {"page-story","first page footer",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of content of probeStoryRange as integer,end of content of probeStoryRange as integer,content of probeStoryRange as text}
set probePhase to "story-availability-and-chain:120: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "story-availability-and-chain:121: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "story-availability-and-chain:122: set end of qualityLayout to {\"page-field\",\"first page footer\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","first page footer",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "story-availability-and-chain:124: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:125: try"
log probePhase
try
set probePhase to "story-availability-and-chain:126: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:128: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:130: try"
log probePhase
try
set probePhase to "story-availability-and-chain:131: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:133: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "story-availability-and-chain:136: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "story-availability-and-chain:137: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:138: try"
log probePhase
try
set probePhase to "story-availability-and-chain:139: set probeStoryRange to get story range boundDoc story type even pages header story"
log probePhase
set probeStoryRange to get story range boundDoc story type even pages header story
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:141: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:143: try"
log probePhase
try
set probePhase to "story-availability-and-chain:144: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:146: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:148: set end of qualityLayout to {\"page-story-available\",\"even pages header\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","even pages header",probeStoryAvailable as boolean}
set probePhase to "story-availability-and-chain:149: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "story-availability-and-chain:150: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "story-availability-and-chain:151: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "story-availability-and-chain:152: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "story-availability-and-chain:153: set end of qualityLayout to {\"page-story\",\"even pages header\",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of co"
log probePhase
set end of qualityLayout to {"page-story","even pages header",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of content of probeStoryRange as integer,end of content of probeStoryRange as integer,content of probeStoryRange as text}
set probePhase to "story-availability-and-chain:154: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "story-availability-and-chain:155: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "story-availability-and-chain:156: set end of qualityLayout to {\"page-field\",\"even pages header\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","even pages header",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "story-availability-and-chain:158: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:159: try"
log probePhase
try
set probePhase to "story-availability-and-chain:160: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:162: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:164: try"
log probePhase
try
set probePhase to "story-availability-and-chain:165: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:167: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "story-availability-and-chain:170: set probeStoryRange to missing value"
log probePhase
set probeStoryRange to missing value
set probePhase to "story-availability-and-chain:171: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:172: try"
log probePhase
try
set probePhase to "story-availability-and-chain:173: set probeStoryRange to get story range boundDoc story type even pages footer story"
log probePhase
set probeStoryRange to get story range boundDoc story type even pages footer story
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:175: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:177: try"
log probePhase
try
set probePhase to "story-availability-and-chain:178: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:180: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:182: set end of qualityLayout to {\"page-story-available\",\"even pages footer\",probeStoryAvailable as boolean}"
log probePhase
set end of qualityLayout to {"page-story-available","even pages footer",probeStoryAvailable as boolean}
set probePhase to "story-availability-and-chain:183: set probeStoryOrdinal to 0"
log probePhase
set probeStoryOrdinal to 0
set probePhase to "story-availability-and-chain:184: repeat while probeStoryAvailable"
log probePhase
repeat while probeStoryAvailable
set probePhase to "story-availability-and-chain:185: set probeStoryOrdinal to probeStoryOrdinal + 1"
log probePhase
set probeStoryOrdinal to probeStoryOrdinal + 1
set probePhase to "story-availability-and-chain:186: if probeStoryOrdinal > 10000 then error \"WPSC_STORY_CHAIN_INVALID\""
log probePhase
if probeStoryOrdinal > 10000 then error "WPSC_STORY_CHAIN_INVALID"
set probePhase to "story-availability-and-chain:187: set end of qualityLayout to {\"page-story\",\"even pages footer\",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of co"
log probePhase
set end of qualityLayout to {"page-story","even pages footer",probeStoryOrdinal as integer,story type of probeStoryRange as text,start of content of probeStoryRange as integer,end of content of probeStoryRange as integer,content of probeStoryRange as text}
set probePhase to "story-availability-and-chain:188: repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange"
log probePhase
repeat with probePageFieldOrdinal from 1 to count fields of probeStoryRange
set probePhase to "story-availability-and-chain:189: set qfield to field probePageFieldOrdinal of probeStoryRange"
log probePhase
set qfield to field probePageFieldOrdinal of probeStoryRange
set probePhase to "story-availability-and-chain:190: set end of qualityLayout to {\"page-field\",\"even pages footer\",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as tex"
log probePhase
set end of qualityLayout to {"page-field","even pages footer",probeStoryOrdinal,probePageFieldOrdinal as integer,field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,start of content of field code of qfield as integer,end of content of field code of qfield as integer,start of content of result range of qfield as integer,end of content of result range of qfield as integer,locked of qfield as boolean}
end repeat
set probePhase to "story-availability-and-chain:192: set probeStoryAvailable to false"
log probePhase
set probeStoryAvailable to false
set probePhase to "story-availability-and-chain:193: try"
log probePhase
try
set probePhase to "story-availability-and-chain:194: set probeStoryRange to next story range of probeStoryRange"
log probePhase
set probeStoryRange to next story range of probeStoryRange
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:196: if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -5941 then error probeStoryError number probeStoryNumber
end try
set probePhase to "story-availability-and-chain:198: try"
log probePhase
try
set probePhase to "story-availability-and-chain:199: set probeStoryAvailable to (probeStoryRange is not missing value)"
log probePhase
set probeStoryAvailable to (probeStoryRange is not missing value)
on error probeStoryError number probeStoryNumber
set probePhase to "story-availability-and-chain:201: if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber"
log probePhase
if probeStoryNumber is not -2753 then error probeStoryError number probeStoryNumber
end try
end repeat
set probePhase to "story-availability-and-chain:204: set probeRows to qualityLayout"
log probePhase
set probeRows to qualityLayout
set nativeRows to {{"read-ok",probeRows}}
on error probeError number probeNumber
if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber
log {probePhase,probeNumber}
set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}
end try
end tell
end timeout
return my jsonRows({"WPSCOMPOSER_WORD_SESSION_OK", nativeRows})
