#import <Foundation/Foundation.h>
#import <AppKit/AppKit.h>
#import <Carbon/Carbon.h>
#import <OSAKit/OSAKit.h>
static OSASendUPP prior; static SRefCon priorRef; static pid_t targetPID;
static OSErr sendBound(const AppleEvent *event,AppleEvent *reply,AESendMode mode,AESendPriority priority,SInt32 timeout,AEIdleUPP idle,AEFilterUPP filter,SRefCon ref) {
 AEDesc address={typeNull,NULL}; OSErr err=AEGetAttributeDesc(event,keyAddressAttr,typeWildCard,&address);
 if(err)return err;
 ProcessSerialNumber psn={0,0};
 if(address.descriptorType==typeProcessSerialNumber && AEGetDescDataSize(&address)==sizeof(psn))AEGetDescData(&address,&psn,sizeof(psn));
 BOOL current=address.descriptorType==typeProcessSerialNumber && psn.highLongOfPSN==0 && psn.lowLongOfPSN==kCurrentProcess;
 AEDisposeDesc(&address);
 if(current)return InvokeOSASendUPP(event,reply,mode,priority,MIN(timeout,600),idle,filter,priorRef,prior);
 if(![[NSRunningApplication runningApplicationWithProcessIdentifier:targetPID].bundleIdentifier isEqual:@"com.microsoft.Excel"])return procNotFound;
 AppleEvent copy={typeNull,NULL}; err=AEDuplicateDesc(event,&copy);
 if(!err)err=AEPutAttributePtr(&copy,keyAddressAttr,typeKernelProcessID,&targetPID,sizeof(targetPID));
 if(!err)err=InvokeOSASendUPP(&copy,reply,mode,priority,MIN(timeout,600),idle,filter,priorRef,prior);
 AEDisposeDesc(&copy);return err;
}
int main(int argc,const char *argv[]) {@autoreleasepool {
 if(argc!=3)return 2;targetPID=atoi(argv[1]);
 NSString *source=[NSString stringWithContentsOfFile:@(argv[2]) encoding:NSUTF8StringEncoding error:nil];if(!source)return 3;
 OSALanguageInstance *instance=[[OSALanguageInstance alloc] initWithLanguage:[OSALanguage languageForName:@"AppleScript"]];
 OSAScript *script=[[OSAScript alloc] initWithSource:source fromURL:nil languageInstance:instance usingStorageOptions:0];
 NSDictionary *error=nil;if(![script compileAndReturnError:&error]){fprintf(stderr,"%s\n",error.description.UTF8String);return 4;}
 ComponentInstance component=instance.componentInstance;OSAGetSendProc(component,&prior,&priorRef);OSASetSendProc(component,NewOSASendUPP(sendBound),0);
 NSAppleEventDescriptor *result=[script executeAndReturnError:&error];
 if(!result){fprintf(stderr,"%s\n",error.description.UTF8String);return 5;}
 puts(result.stringValue.UTF8String);return 0;
}}
