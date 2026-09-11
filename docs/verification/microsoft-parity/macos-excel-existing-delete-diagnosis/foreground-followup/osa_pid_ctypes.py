"""Throwaway comparison to the compiled diagnostic; not a product helper."""
import ctypes as C
import sys
from pathlib import Path
U=C.c_uint32; I=C.c_int32; P=C.c_void_p
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
getsize=fn('AEGetDescDataSize',[D],I)
getdata=fn('AEGetDescData',[D,P,I],C.c_int16)
compile_=fn('OSACompile',[P,D,I,C.POINTER(U)])
execute=fn('OSAExecute',[P,U,U,I,C.POINTER(U)])
display=fn('OSADisplay',[P,U,U,I,D])
script_error=fn('OSAScriptError',[P,U,U,D])
osa_dispose=fn('OSADispose',[P,U])
CALLBACK=C.CFUNCTYPE(C.c_int16,D,D,I,C.c_int16,I,P,P,C.c_ssize_t)
getsend=fn('OSAGetSendProc',[P,C.POINTER(P),C.POINTER(C.c_ssize_t)])
setsend=fn('OSASetSendProc',[P,CALLBACK,C.c_ssize_t])
def check(err):
 if err:raise RuntimeError(err)
def text(desc):
 n=getsize(C.byref(desc));buf=C.create_string_buffer(n+1);check(getdata(C.byref(desc),buf,n));return buf.raw[:n].decode('utf8')
pid=I(int(sys.argv[1]));component=open_component(code('osa '),code('ascr'))
source=Desc();compiled=U();result=U();output=Desc();old=P();ref=C.c_ssize_t();
raw=Path(sys.argv[2]).read_bytes();check(create(code('utf8'),raw,len(raw),C.byref(source)))
check(compile_(component,C.byref(source),0,C.byref(compiled)))
check(getsend(component,C.byref(old),C.byref(ref)));oldsend=CALLBACK(old.value)
@CALLBACK
def send(event,reply,mode,priority,timeout,idle,filter_,unused):
 copy=Desc()
 try:
  check(duplicate(event,C.byref(copy)))
  check(putattr(C.byref(copy),code('addr'),code('kpid'),C.byref(pid),C.sizeof(pid)))
  return oldsend(C.byref(copy),reply,mode,priority,min(timeout,600),idle,filter_,ref.value)
 except Exception as exc:
  print(str(exc),file=sys.stderr);return -50
 finally:dispose(C.byref(copy))
check(setsend(component,send,0))
err=execute(component,compiled,0,0,C.byref(result))
if err:
 script_error(component,code('errs'),code('utf8'),C.byref(output));print('OSA error',err,text(output),file=sys.stderr)
else:
 check(display(component,result,code('utf8'),0,C.byref(output)));print(text(output))
for d in (output,source):dispose(C.byref(d))
for x in (result,compiled):
 if x.value:osa_dispose(component,x)
close_component(component)
sys.exit(1 if err else 0)
