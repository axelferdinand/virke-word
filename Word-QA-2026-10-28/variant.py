from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/virke-build'))
from revise_designer_feedback import E,q,NS,child,remove,order,xml
root=Path(__file__).resolve().parents[1]
source=root/'Virke-revidert-2026-10-27/02_Balance_scorecard.docx'
with ZipFile(source) as z:data={n:z.read(n) for n in z.namelist()}
for part in ['word/styles.xml','word/glossary/styles.xml']:
 s=E.fromstring(data[part]);normal=s.find('w:style[@w:styleId="Normal"]',NS)
 dd=child(s,'docDefaults')
 for prop,container in [('rPr','rPrDefault'),('pPr','pPrDefault')]:
  target=child(child(dd,container),prop)
  for n in list(child(normal,prop)):
   old=target.find(n.tag)
   if old is not None:target.remove(old)
   target.append(deepcopy(n))
  remove(normal,prop)
 for sid in ['Virke1','Virke2']:
  st=s.find('w:style[@w:styleId="'+sid+'"]',NS)
  mar=child(child(st,'tblPr'),'tblCellMar')
  for k in ['top','bottom','left','right']:child(mar,k,w=120,type='dxa')
  child(child(st,'tcPr'),'vAlign',val='center')
  for cond in st.findall(q('tblStylePr')):child(child(cond,'tcPr'),'vAlign',val='center')
 order(s);data[part]=xml(s)
for part in ['word/document.xml','word/glossary/document.xml']:
 d=E.fromstring(data[part])
 for t in d.iter(q('tbl')):
  st=t.find('w:tblPr/w:tblStyle',NS)
  if st is None or st.get(q('val')) not in ['Virke1','Virke2']:continue
  for mar in t.findall('w:tblPr/w:tblCellMar',NS)+t.findall('w:tr/w:tc/w:tcPr/w:tcMar',NS):
   for k in ['top','bottom','left','right']:child(mar,k,w=120,type='dxa')
  for cp in t.findall('w:tr/w:tc/w:tcPr',NS):child(cp,'vAlign',val='center')
 order(d);data[part]=xml(d)
out=root/'Word-QA-2026-10-28/testkopier/28okt-variant-A.docx'
with ZipFile(out,'w',ZIP_DEFLATED) as z:
 for n,b in data.items():z.writestr(n,b)
print(out)
