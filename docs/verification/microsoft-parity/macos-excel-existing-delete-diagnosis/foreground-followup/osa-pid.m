#import <Foundation/Foundation.h>
#import <AppKit/AppKit.h>
#import <Carbon/Carbon.h>
static OSASendUPP previous; static SRefCon previousRef; static pid_t ownedPID=78330;
static OSErr sendOwned(const AppleEvent *event,AppleEvent *reply,AESendMode mode,AESendPriority priority,SInt32 timeout,AEIdleUPP idle,AEFilterUPP filter,SRefCon ref) {
 if (![[NSRunningApplication runningApplicationWithProcessIdentifier:ownedPID].bundleIdentifier isEqual:@"com.microsoft.Excel"]) return procNotFound;
 AppleEvent copy={typeNull,NULL}; OSErr err=AEDuplicateDesc(event,&copy); if(err)return err;
 err=AEPutAttributePtr(&copy,keyAddressAttr,typeKernelProcessID,&ownedPID,sizeof(ownedPID));
 fprintf(stderr,"send %s\n",[[NSAppleEventDescriptor alloc] initWithAEDescNoCopy:&copy].description.UTF8String);
 // The descriptor above owns the copied descriptor. Keep a second copy for send.
 AppleEvent addressed={typeNull,NULL}; err=AEDuplicateDesc(event,&addressed);
 if (!err) err=AEPutAttributePtr(&addressed,keyAddressAttr,typeKernelProcessID,&ownedPID,sizeof(ownedPID));
 if(!err) err=InvokeOSASendUPP(&addressed,reply,mode,priority,MIN(timeout,600),idle,filter,previousRef,previous);
 AEDisposeDesc(&addressed); return err;
}
int main(int argc,const char*argv[]) { @autoreleasepool {
 if(argc!=2)return 2;
 NSString *source=[NSString stringWithContentsOfFile:@(argv[1]) encoding:NSUTF8StringEncoding error:nil];
 if(!source)return 3;
 ComponentInstance component=OpenDefaultComponent(kOSAComponentType,kAppleScriptSubtype);
 OSAID script=kOSANullScript,result=kOSANullScript;
 NSAppleEventDescriptor *text=[NSAppleEventDescriptor descriptorWithString:source];
 OSAError err=OSACompile(component,text.aeDesc,kOSAModeNull,&script);
 if(!err)err=OSAGetSendProc(component,&previous,&previousRef);
 if(!err)err=OSASetSendProc(component,NewOSASendUPP(sendOwned),0);
 if(!err)err=OSAExecute(component,script,kOSANullScript,kOSAModeNull,&result);
 AEDesc out={typeNull,NULL};
 if(err) {OSAScriptError(component,kOSAErrorMessage,typeUTF8Text,&out);fprintf(stderr,"OSA error %d\n",(int)err);}
 else OSADisplay(component,result,typeUTF8Text,kOSAModeNull,&out);
 Size n=AEGetDescDataSize(&out); char*buf=calloc(n+1,1); AEGetDescData(&out,buf,n);puts(buf);free(buf);AEDisposeDesc(&out);
 if(result)OSADispose(component,result);if(script)OSADispose(component,script);CloseComponent(component);return err?1:0;
} }
