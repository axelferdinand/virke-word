"""Repair the table style cascade, preserving original package parts.

Normal inherits document defaults so table styles can own size, font and spacing.
Word/Clan needs an optical top/bottom correction compared with Figma's CSS line
box. 13.5 pt vertical padding reproduces the 35 px minimum row height in new
tables too (Word ignores style-level minimum row heights). Native review gates release.
"""
from pathlib import Path
from copy import deepcopy
from zipfile import ZipFile, ZIP_DEFLATED
import json, hashlib
from revise_designer_feedback import E, q, NS, child, remove, order, xml, txt

ROOT=Path(__file__).resolve().parents[2]
SOURCE=ROOT/'Virke-revidert-2026-10-27'
DEST=ROOT/'Word-QA-2026-10-28/kandidater'

def style(r,sid):return r.find('w:style[@w:styleId="'+sid+'"]',NS)

def margins(m):
    for side,value in [('top',105),('left',120),('bottom',165),('right',120)]:
        child(m,side,w=value,type='dxa')

def styles(r):
    normal=style(r,'Normal');dd=child(r,'docDefaults')
    # Explicit Normal values outrank table-style values. Move them to the
    # lower-priority document defaults; body paragraphs retain their appearance.
    for prop,container in [('rPr','rPrDefault'),('pPr','pPrDefault')]:
        dest=child(child(dd,container),prop)
        for n in list(child(normal,prop)):
            old=dest.find(n.tag)
            if old is not None:dest.remove(old)
            dest.append(deepcopy(n))
        remove(normal,prop)
    for sid in ['Virke1','Virke2']:
        s=style(r,sid)
        margins(child(child(s,'tblPr'),'tblCellMar'))
        child(child(s,'tcPr'),'vAlign',val='center')
        if sid=='Virke1':
            child(child(s,'pPr'),'spacing',before=0,after=0,line=240,lineRule='atLeast')
            for key in ['sz','szCs']:child(child(s,'rPr'),key,val=15)
        for cond in s.findall(q('tblStylePr')):
            child(child(cond,'tcPr'),'vAlign',val='center')
    # These paragraph styles must not compete with the table header condition.
    for sid in ['VirkeTabelltekst','VirkeTabelloverskrift']:
        rp=child(style(r,sid),'rPr');remove(rp,'b','bCs')
    large=deepcopy(style(r,'VirkeTabelltekst'))
    large.set(q('styleId'),'VirkeTabellLangtekst')
    child(large,'name',val='Virke tabell langtekst')
    for key in ['sz','szCs']:child(child(large,'rPr'),key,val=18)
    child(child(large,'pPr'),'spacing',before=0,after=0,line=288,lineRule='atLeast')
    r.append(large)
    order(r)

def body(r):
    for t in r.iter(q('tbl')):
        st=t.find('w:tblPr/w:tblStyle',NS)
        sid=st.get(q('val')) if st is not None else ''
        pstyles={x.get(q('val')) for x in t.findall('.//w:pStyle',NS)}
        data=sid in ['Virke1','Virke2']
        form=pstyles=={'VirkeTabelltekst'}
        if not (data or form):continue
        margins(child(child(t,'tblPr'),'tblCellMar'))
        rows=t.findall(q('tr'))
        large_cells=sid=='Virke1' and len(rows)>1 and [txt(c) for c in rows[0].findall(q('tc'))]==['Avstemninger','OK','I/A','Oppfølgingsansvar','Merknad']
        for i,tr in enumerate(rows):
            if sid=='Virke1':remove(child(tr,'trPr'),'trHeight')
            for tc in tr.findall(q('tc')):
                cp=child(tc,'tcPr');margins(child(cp,'tcMar'))
                child(cp,'vAlign',val='center')
                if sid=='Virke1' and i==0:
                    # Conditional table style is the single owner of header
                    # weight. Clear run/paragraph-mark/field typing overrides.
                    for rp in tc.findall('.//w:rPr',NS):remove(rp,'b','bCs')
        # Text-heavy body rows use top alignment, as in the Figma multiline
        # example. Header cells remain vertically centred.
        if sid=='Virke1':
            for tr in rows[1:]:
                long_row=any(len(txt(c))>70 or len(c.findall('w:p',NS))>1 for c in tr.findall(q('tc')))
                if long_row:
                    for tc in tr.findall(q('tc')):child(child(tc,'tcPr'),'vAlign',val='top')
        if large_cells:
            for tr in rows[1:]:
                first=tr.find(q('tc'))
                for i,p in enumerate(first.findall(q('p'))):
                    child(child(p,'pPr'),'pStyle',val='VirkeTabellLangtekst')
                    if i==0:child(child(p,'pPr'),'spacing',after=180)
    order(r)

def build(source,dest):
    with ZipFile(source) as z:original={n:z.read(n) for n in z.namelist()}
    data=dict(original)
    for part in ['word/styles.xml','word/glossary/styles.xml']:
        s=E.fromstring(data[part]);styles(s);data[part]=xml(s)
    for part in ['word/document.xml','word/glossary/document.xml']:
        s=E.fromstring(data[part]);body(s)
        if part=='word/document.xml' and source.name.startswith('04_'):
            # Keep the one-page example after correcting row height. The last
            # row is wholly empty placeholders; editors can add rows with Tab.
            table=s.findall('.//w:tbl',NS)[-1];last=table.findall(q('tr'))[-1]
            assert all(txt(c)=='[Fyll ut]' for c in last.findall(q('tc')))
            table.remove(last)
        data[part]=xml(s)
    for n,b in data.items():
        if n.endswith(('.xml','.rels')):E.fromstring(b)
        if any(x in n for x in ['/charts/','/embeddings/','/media/','/theme/']):assert b==original[n],n
    dest.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(dest,'w',ZIP_DEFLATED) as z:
        for n,b in data.items():z.writestr(n,b)
    return {'file':str(dest.relative_to(ROOT)), 'sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),
            'changed_parts':[n for n in data if data[n]!=original[n]]}

if __name__=='__main__':
    result=[build(p,DEST/p.name) for p in sorted(SOURCE.iterdir()) if p.suffix in ['.docx','.dotx'] and 'brukerveiledning' not in p.name]
    (DEST.parent/'candidate-build.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
    print(f'Built {len(result)} candidates; native Word verification pending.')
