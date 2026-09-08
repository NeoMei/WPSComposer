import sys, os, subprocess, time, json
from pathlib import Path
root=Path.cwd()
label=sys.argv[1]
args=sys.argv[2:]
out=root/'build/msoffice-production/logs'
env=os.environ.copy()
env['PYTHONPATH']=os.pathsep.join([str(root/'build/msoffice-production/audit-hooks'),str(root/'build/msoffice-production/deps'),'D:/MyDocuments/codexprojects/WpsComposer/build/msoffice-spike/audit-pytest-deps',str(root)])
env['WPSCOMPOSER_ACCEPTANCE_TRACE']=str(root/'build/msoffice-production/traces'/label)
started=time.time()
with (out/(label+'.stdout.log')).open('xb') as stdout, (out/(label+'.stderr.log')).open('xb') as stderr:
    result=subprocess.run([sys.executable]+args,env=env,stdout=stdout,stderr=stderr)
record={'command':[sys.executable]+args,'cwd':str(root),'exit_code':result.returncode,'started_unix':started,'ended_unix':time.time(),'observation_hook':'audit-hooks/sitecustomize.py'}
(out/(label+'-command.json')).write_text(json.dumps(record,indent=2),encoding='utf-8')
print(json.dumps(record))
raise SystemExit(result.returncode)
