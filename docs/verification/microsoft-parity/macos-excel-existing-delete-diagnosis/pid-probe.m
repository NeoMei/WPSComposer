#import <Foundation/Foundation.h>
#import <ScriptingBridge/ScriptingBridge.h>
#import <AppKit/AppKit.h>
static void check(BOOL ok, NSString *message) { if (!ok) @throw [NSException exceptionWithName:@"ProbeFailed" reason:message userInfo:nil]; }
static NSArray *inventory(SBApplication *app) {
 NSMutableArray *out=[NSMutableArray array];
 for (SBObject *b in [app valueForKey:@"workbooks"]) {
  NSMutableArray *sheets=[NSMutableArray array];
  for (SBObject *ws in [b valueForKey:@"worksheets"]) [sheets addObject:[ws valueForKey:@"name"]];
  [out addObject:@{@"name":[b valueForKey:@"name"],@"path":[b valueForKey:@"fullName"],@"saved":[b valueForKey:@"saved"],@"sheets":sheets}];
 }
 return out;
}
int main(int argc,const char *argv[]) { @autoreleasepool { @try {
 check(argc>=3,@"pid mode required");
 pid_t pid=atoi(argv[1]);
 NSRunningApplication *run=[NSRunningApplication runningApplicationWithProcessIdentifier:pid];
 check([[run bundleIdentifier] isEqual:@"com.microsoft.Excel"],@"Excel PID mismatch");
 SBApplication *app=[SBApplication applicationWithProcessIdentifier:pid]; app.timeout=120;
 NSString *mode=@(argv[2]);
 NSArray *before=inventory(app);
 if ([mode isEqual:@"open"]) {
  check(pid==99513 && argc==4,@"Owned PID and path required");
  check(before.count==1 && [before[0][@"name"] isEqual:@"工作簿1"] && [before[0][@"saved"] boolValue],@"Unexpected owned instance inventory");
  NSString *path=@(argv[3]);
  check([path containsString:@"excel-existing-delete-diagnosis/"] && [[NSFileManager defaultManager] fileExistsAtPath:path],@"Owned fixture path required");
  NSAppleEventDescriptor *event=[NSAppleEventDescriptor appleEventWithEventClass:'smXL' eventID:'1169' targetDescriptor:[NSAppleEventDescriptor descriptorWithProcessIdentifier:pid] returnID:kAutoGenerateReturnID transactionID:kAnyTransactionID];
  [event setParamDescriptor:[NSAppleEventDescriptor descriptorWithString:CFBridgingRelease(CFURLCopyFileSystemPath((__bridge CFURLRef)[NSURL fileURLWithPath:path],kCFURLHFSPathStyle))] forKeyword:'WbFN'];
  [event setParamDescriptor:[NSAppleEventDescriptor descriptorWithEnumCode:0x02b60000] forKeyword:'XOul'];
  [event setParamDescriptor:[NSAppleEventDescriptor descriptorWithBoolean:YES] forKeyword:'XoEd'];
  [event setParamDescriptor:[NSAppleEventDescriptor descriptorWithBoolean:NO] forKeyword:'5245'];
  NSError *error=nil;
  NSAppleEventDescriptor *reply=[event sendEventWithOptions:NSAppleEventSendWaitForReply timeout:10 error:&error];
  fprintf(stderr,"reply: %s; error: %s\n",reply.description.UTF8String,error.description.UTF8String);
  check(reply!=nil && error==nil && [reply paramDescriptorForKeyword:'errn'].int32Value==0,@"Raw open failed");

 }
 NSDictionary *result=@{@"pid":@(pid),@"alerts":[app valueForKey:@"displayAlerts"],@"before":before,@"after":inventory(app)};
 NSData *data=[NSJSONSerialization dataWithJSONObject:result options:NSJSONWritingPrettyPrinted|NSJSONWritingSortedKeys error:nil];
 puts([[NSString alloc] initWithData:data encoding:NSUTF8StringEncoding].UTF8String);
 return 0;
 } @catch(NSException *e) { fprintf(stderr,"%s: %s\n",e.name.UTF8String,e.reason.UTF8String); return 1; } } }
