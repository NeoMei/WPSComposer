"""Throwaway comparison to the compiled diagnostic; not a product helper."""
import ctypes as C
import sys
from pathlib import Path
U=C.c_uint32; I=C.c_int32; P=C.c_void_p
C.CDLL('/System/Library/Frameworks/OSAKit.framework/OSAKit')
objc=C.CDLL('/usr/lib/libobjc.A.dylib')
objc.objc_getClass.argtypes=[C.c_char_p];objc.objc_getClass.restype=P
objc.sel_registerName.argtypes=[C.c_char_p];objc.sel_registerName.restype=P
objc.objc_msgSend.argtypes=[P,P];objc.objc_msgSend.restype=P
def msg(receiver,name,args=(),types=(),result=P):
 address=C.cast(objc.objc_msgSend,P).value
 call=C.CFUNCTYPE(result,P,P,*types)(address)
 return call(receiver,objc.sel_registerName(name.encode()),*args)
def cls(name):return objc.objc_getClass(name.encode())
def string(value):return msg(cls('NSString'),'stringWithUTF8String:',(value.encode(),),(C.c_char_p,))
pool=msg(cls('NSAutoreleasePool'),'new')

class Desc(C.Structure): _fields_=[('kind',U),('data',P)]
D=C.POINTER(Desc)
def code(s):return int.from_bytes(s.encode('ascii'),'big')
lib=C.CDLL('/System/Library/Frameworks/Carbon.framework/Carbon')
def fn(name,args,result=I):
 f=getattr(lib,name);f.argtypes=args;f.restype=result;return f
open_component=fn('OpenDefaultComponent',[U,U],P)
close_component=fn('CloseComponent',[P])
create=fn('AECreateDesc',[U,P,I,D],C.c_int16)
dispose=fn('AEDisposeDesc',[D],C.c_int16)
duplicate=fn('AEDuplicateDesc',[D,D],C.c_int16)
putattr=fn('AEPutAttributePtr',[D,U,U,P,I],C.c_int16)
getattr_=fn('AEGetAttributeDesc',[D,U,U,D],C.c_int16)
getsize=fn('AEGetDescDataSize',[D],I)
getdata=fn('AEGetDescData',[D,P,I],C.c_int16)
compile_=fn('OSACompile',[P,D,I,C.POINTER(U)])
execute=fn('OSAExecute',[P,U,U,I,C.POINTER(U)])
display=fn('OSACoerceToDesc',[P,U,U,I,D])
script_error=fn('OSAScriptError',[P,U,U,D])
osa_dispose=fn('OSADispose',[P,U])
CALLBACK=C.CFUNCTYPE(C.c_int16,D,D,I,C.c_int16,I,P,P,C.c_ssize_t)
getsend=fn('OSAGetSendProc',[P,C.POINTER(P),C.POINTER(C.c_ssize_t)])
setsend=fn('OSASetSendProc',[P,CALLBACK,C.c_ssize_t])
def check(err):
 if err:raise RuntimeError(err)
def text(desc):
 n=getsize(C.byref(desc));buf=C.create_string_buffer(n+1);check(getdata(C.byref(desc),buf,n));return buf.raw[:n].decode('utf8')
pid=I(int(sys.argv[1]));old=P();ref=C.c_ssize_t()
language=msg(cls('OSALanguage'),'languageForName:',(string('AppleScript'),),(P,))
instance=msg(msg(cls('OSALanguageInstance'),'alloc'),'initWithLanguage:',(language,),(P,))
source=string(Path(sys.argv[2]).read_text())
script=msg(msg(cls('OSAScript'),'alloc'),'initWithSource:fromURL:languageInstance:usingStorageOptions:',(source,None,instance,0),(P,P,P,C.c_ulong))
error=P()
if not msg(script,'compileAndReturnError:',(C.byref(error),),(C.POINTER(P),),C.c_bool):raise RuntimeError('Compile failed')
component=msg(instance,'componentInstance')
check(getsend(component,C.byref(old),C.byref(ref)));oldsend=CALLBACK(old.value)
@CALLBACK
def send(event,reply,mode,priority,timeout,idle,filter_,unused):
 copy=Desc()
 try:
  address=Desc()
  check(getattr_(event,code('addr'),code('****'),C.byref(address)))
  n=getsize(C.byref(address));buf=C.create_string_buffer(n);getdata(C.byref(address),buf,n)

  print('TARGET',address.kind.to_bytes(4,'big'),buf.raw.hex(),file=sys.stderr,flush=True)
  is_current = address.kind == code('psn ') and buf.raw == bytes.fromhex('0000000002000000')
  dispose(C.byref(address))
  if is_current:
   return oldsend(event,reply,mode,priority,min(timeout,600),idle,filter_,ref.value)
  check(duplicate(event,C.byref(copy)))
  check(putattr(C.byref(copy),code('addr'),code('kpid'),C.byref(pid),C.sizeof(pid)))
  return oldsend(C.byref(copy),reply,mode,priority,min(timeout,600),idle,filter_,ref.value)
 except Exception as exc:
  print(str(exc),file=sys.stderr);return -50
 finally:dispose(C.byref(copy))
check(setsend(component,send,0))
result=msg(script,'executeAndReturnError:',(C.byref(error),),(C.POINTER(P),))
def ns_text(obj):return msg(obj,'UTF8String',result=C.c_char_p).decode()
if not result:
 print(ns_text(msg(error.value,'description')),file=sys.stderr)
 sys.exit(1)
print(ns_text(msg(result,'stringValue')))
