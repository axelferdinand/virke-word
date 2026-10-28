import sys,json,shutil
from pathlib import Path
from zipfile import ZipFile
sys.path.insert(0,'scripts/virke-build')
import revise_feedback_2026_10_27 as r
changed=[]
for src in sorted(p for p in r.SOURCE.iterdir() if p.suffix in ['.docx','.dotx'] and 'brukerveiledning' not in p.name):
 tmp=r.QA/'gap-build'/src.name
 r.build_one(src,tmp)
 target=r.DEST/src.name
 with ZipFile(tmp) as a,ZipFile(target) as b:
  dif=[n for n in a.namelist() if a.read(n)!=b.read(n)]
 if dif:
  shutil.copy2(target,r.QA/'gap-build'/('before-'+src.name))
  shutil.copy2(tmp,target)
  shutil.copy2(tmp,r.QA/'sluttkontroll'/('Slutt-'+src.name))
  changed.append({'file':src.name,'parts':dif})
print(json.dumps(changed,indent=2))
(r.QA/'gap-correction.json').write_text(json.dumps(changed,indent=2))
