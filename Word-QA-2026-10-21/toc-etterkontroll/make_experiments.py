from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
from copy import deepcopy
R=Path(__file__).resolve().parent
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';NS={'w':W};q=lambda x:'{'+W+'}'+x
with ZipFile(R.parents[1]/'Virke-revidert-2026-10-21/07_Innholdsfortegnelse_og_avtale.docx') as z:base={n:z.read(n) for n in z.namelist()}
def sub(p,t,**a):return E.SubElement(p,q(t),{q(k):str(v) for k,v in a.items()})
for mode in ['black','auto']:
 parts=base.copy();s=E.fromstring(parts['word/styles.xml']);d=E.fromstring(parts['word/document.xml']);b=d.find(q('body'))
 toc=deepcopy(b.find(q('sdt')));sect=deepcopy(b.find(q('sectPr')));b.clear();b.append(toc)
 s.find('w:style[@w:styleId="INNH1"]/w:rPr/w:color',NS).set(q('val'),'616161')
 if mode=='auto':s.find('w:style[@w:styleId="Overskrift1"]/w:rPr/w:color',NS).set(q('val'),'auto')
 cs=sub(s,'style',type='character',customStyle='1',styleId='QAHeadingColor');sub(cs,'name',val='QA Heading Color');rp=sub(cs,'rPr');sub(rp,'color',val='000000')
 first=toc.xpath('.//w:p[w:pPr/w:pStyle[@w:val="INNH1"]]',namespaces=NS)[0]
 pp=first.find(q('pPr'));bd=sub(pp,'pBdr');sub(bd,'top',val='nil');sub(bd,'between',val='nil')
 # Clear cached hyperlinks so full generation cannot preserve old title colors.
 for h in list(toc.iter(q('hyperlink'))):h.getparent().remove(h)
 tests=[('A Vanlig Overskrift 1',None),('B Direkte svart 000000','000000'),('C Direkte nesten svart 010101','010101'),('D Tegnstil svart','style'),('E Direkte automatisk','auto'),('F Ny vanlig Overskrift 1',None)]
 for i,(title,color) in enumerate(tests):
  p=sub(b,'p');pp=sub(p,'pPr');sub(pp,'pStyle',val='Overskrift1')
  if i in (0,3):sub(pp,'pageBreakBefore')
  run=sub(p,'r')
  if color:
   rp=sub(run,'rPr');sub(rp,'rStyle' if color=='style' else 'color',val='QAHeadingColor' if color=='style' else color)
  sub(run,'t').text=title
  for level,label in [(2,'Underkapittel'),(3,'Detaljer')]:
   p=sub(b,'p');pp=sub(p,'pPr');sub(pp,'pStyle',val='Overskrift'+str(level));sub(sub(p,'r'),'t').text=label+' '+str(i+1)
 b.append(sect)
 for n,x in [('word/styles.xml',s),('word/document.xml',d)]:parts[n]=E.tostring(x,encoding='UTF-8',xml_declaration=True,standalone=True)
 with ZipFile(R/f'TOC-test-{mode}.docx','w',ZIP_DEFLATED) as z:
  for n,v in parts.items():z.writestr(n,v)
print('Created two isolated TOC test documents.')
