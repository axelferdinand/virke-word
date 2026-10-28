"""Preserve the 21 October packages and apply Emils 27 October changes.

Builds candidates only. Native Word tests and guide regeneration are separate
delivery gates; a successful build is not a delivery approval.
"""
from pathlib import Path
from copy import deepcopy
from zipfile import ZipFile, ZIP_DEFLATED
import argparse, hashlib, json, uuid
from lxml import etree as E
from revise_designer_feedback import W, NS, q, child, remove, order, xml, txt, block

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / 'Virke-med-komponentguide-2026-10-21'
QA = ROOT / 'Word-QA-2026-10-27'
DEST = QA / 'kandidater'

RIBBON = '''<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon><tabs>
<tab idMso="TabHome"><group id="VirkeHomeBlocks" label="Virke" insertBeforeMso="GroupClipboard"><button idMso="AutoManager" label="Virke-blokker" imageMso="AutoTextGallery" size="large" screentip="Tabeller, signatur og innholdsfortegnelse" supertip="Velg fanen Autotekst og Virke-masteren under Søk i. Velg en blokk og Sett inn."/></group></tab>
<tab id="VirkeTab" label="Virke" insertAfterMso="TabHome">
<group id="VirkeInsert" label="Tabeller, skjema og signatur"><button idMso="AutoManager" label="Virke-blokker" imageMso="AutoTextGallery" size="large" screentip="Åpne Autotekst" supertip="Velg fanen Autotekst og Virke-masteren under Søk i. Velg en blokk og Sett inn."/></group>
<group id="VirkeContents" label="Innholdsfortegnelse"><button idMso="TableOfContentsDialog" label="Sett inn innholdsfortegnelse" size="large" screentip="Innholdsfortegnelse fra malens stiler" supertip="Bruk formatet Fra mal og tre nivåer. Malens ferdige innholdsfortegnelse finnes også under Virke-blokker."/><button idMso="FieldsUpdate" label="Oppdater felt" size="large"/></group>
</tab></tabs></ribbon></customUI>'''

def style(root, sid):
    return root.find('w:style[@w:styleId="'+sid+'"]', NS)

def revise_styles(root):
    for sid, before, after in [('Overskrift1',450,270),('Overskrift3',360,180)]:
        child(child(style(root,sid),'pPr'),'spacing',before=before,after=after)
    # The conditional table style owns header bold, so native Insert Table works.
    # A second bold toggle in the paragraph style would cancel that bold.
    header=style(root,'VirkeTabelloverskrift')
    remove(child(header,'rPr'),'b','bCs')
    for sid in ['Virke1','Virke2']:
        s=style(root,sid)
        mar=child(child(s,'tblPr'),'tblCellMar')
        child(mar,'top',w=180,type='dxa'); child(mar,'bottom',w=60,type='dxa')
        child(child(s,'tcPr'),'vAlign',val='top')
        if sid=='Virke1':
            cond=s.find('w:tblStylePr[@w:type="firstRow"]',NS)
            rp=child(cond,'rPr'); child(rp,'b',val=1);child(rp,'bCs',val=1)
    # Built-in List Bullet must expose the branded existing list too.
    bullet=style(root,'Punktliste')
    remove(bullet,'semiHidden');child(bullet,'qFormat');child(bullet,'uiPriority',val=5)
    child(bullet,'next',val='Punktliste')
    pp=child(bullet,'pPr');child(pp,'spacing',before=0,after=120,line=288,lineRule='atLeast')
    child(pp,'ind',left=340,hanging=340)
    # Word drops the final AutoText paragraph mark during insertion. Keep a
    # dedicated spacer BEFORE a separate Normal continuation paragraph.
    for sid, name, before in [('VirkeEtterTabell','Virke etter tabell',360),
                              ('VirkeEtterModul','Virke etter modul',180)]:
        s=E.SubElement(root,q('style'),{q('type'):'paragraph',q('customStyle'):'1',q('styleId'):sid})
        child(s,'name',val=name);child(s,'basedOn',val='Normal');child(s,'next',val='Normal')
        child(s,'semiHidden');child(s,'uiPriority',val=99)
        child(child(s,'pPr'),'spacing',before=0,after=0,line=before,lineRule='exact')
        child(child(s,'rPr'),'sz',val=2);child(child(s,'rPr'),'szCs',val=2)
    order(root)

def revise_numbering(root):
    # Keep identities and numbered lists; adjust only the existing built-in
    # List Bullet definition. Native autoformat behavior still needs testing.
    num=root.find('w:num[@w:numId="5"]',NS)
    aid=num.find(q('abstractNumId')).get(q('val'))
    a=root.find('w:abstractNum[@w:abstractNumId="'+aid+'"]',NS)
    lvl=a.find(q('lvl'));rp=child(lvl,'rPr')
    child(rp,'color',val='C9007F');child(rp,'sz',val=18)
    child(child(lvl,'pPr'),'ind',left=340,hanging=340)
    order(root)

def revise_body(root):
    for p in root.iter(q('p')):
        ps=p.find('w:pPr/w:pStyle',NS)
        if ps is not None and ps.get(q('val')) in ['Overskrift1','Overskrift3']:
            spacing=p.find('w:pPr/w:spacing',NS)
            if spacing is not None:
                for attr in ['before','after','beforeLines','afterLines','beforeAutospacing','afterAutospacing']:
                    spacing.attrib.pop(q(attr),None)
    for t in root.iter(q('tbl')):
        sid=t.find('w:tblPr/w:tblStyle',NS)
        is_data=sid is not None and sid.get(q('val')) in ['Virke1','Virke2']
        # Form fields have no data-table style but use the same 7.5 pt cell text.
        pstyles={p.get(q('val')) for p in t.findall('.//w:pStyle',NS)}
        is_form=not is_data and pstyles=={'VirkeTabelltekst'}
        is_summary='VirkeOppsummering' in pstyles
        if not (is_data or is_form or is_summary):continue
        mar=child(child(t,'tblPr'),'tblCellMar')
        if not is_summary:
            child(mar,'top',w=180,type='dxa');child(mar,'bottom',w=60,type='dxa')
        for tc in t.findall('w:tr/w:tc',NS):
            cp=child(tc,'tcPr');child(cp,'vAlign',val='top')
            cm=child(cp,'tcMar')
            top=cm.find(q('top'))
            new_top=60 if is_summary and top is not None and top.get(q('w'))=='0' else 180
            child(cm,'top',w=new_top,type='dxa');child(cm,'bottom',w=60,type='dxa')
    order(root)

def table_kind(t):
    sid=t.find('w:tblPr/w:tblStyle',NS)
    pstyles={p.get(q('val')) for p in t.findall('.//w:pStyle',NS)}
    if pstyles.intersection({'VirkeSignatur','VirkeOppsummering'}):return 'module'
    if (sid is not None and sid.get(q('val')) in ['Virke1','Virke2']) or pstyles=={'VirkeTabelltekst'}:return 'table'
    return None

def gap_paragraph(p,kind):
    pp=child(p,'pPr')
    remove(pp,'spacing','rPr','keepNext','keepLines')
    child(pp,'pStyle',val='VirkeEtterTabell' if kind=='table' else 'VirkeEtterModul')

def blank_spacer(p):
    if p is None or p.tag!=q('p') or txt(p):return False
    if p.find('w:pPr/w:sectPr',NS) is not None:return False
    return not any(n.tag in {q('drawing'),q('pict'),q('br'),q('fldChar'),q('fldSimple'),q('sdt'),q('bookmarkStart'),q('bookmarkEnd')} for n in p.iter())

def revise_gaps(root,glossary=False):
    """Keep inline tables in normal flow; never convert to floating tables.

    Glossary blocks end in a spacer and editable continuation paragraph. In populated
    examples, transfer its spacing to the actual following paragraph so a
    technical blank line cannot double the gap.
    """
    for t in list(root.iter(q('tbl'))):
        kind=table_kind(t)
        if kind is None:continue
        parent=t.getparent()
        if parent.tag not in [q('body'),q('docPartBody'),q('sdtContent')]:continue
        prev=t.getprevious()
        if blank_spacer(prev) and prev.getprevious() is not None and prev.getprevious().tag==q('p'):
            prior=prev.getprevious();parent.remove(prev);prev=prior
        if prev is not None and prev.tag==q('p'):
            pp=child(prev,'pPr')
            if blank_spacer(prev):
                # Existing dedicated spacer before a block.
                # Between adjacent tables, preserve the preceding table's
                # 18 pt after-gap instead of replacing it with 9 pt before.
                previous_block=prev.getprevious()
                line=360 if previous_block is not None and previous_block.tag==q('tbl') and table_kind(previous_block)=='table' else 180
                child(pp,'spacing',before=0,after=0,line=line,lineRule='exact')
            else:child(pp,'spacing',after=180)
        tail=t.getnext()
        if tail is None or tail.tag!=q('p'):
            tail=E.Element(q('p'));parent.insert(parent.index(t)+1,tail)
        if glossary:
            if not txt(tail):
                gap_paragraph(tail,kind)
                if tail.getnext() is None:
                    continuation=E.Element(q('p'));child(child(continuation,'pPr'),'pStyle',val='Normal')
                    parent.append(continuation)
            else:child(child(tail,'pPr'),'spacing',before=360 if kind=='table' else 180)
        else:
            while blank_spacer(tail) and tail.getnext() is not None and tail.getnext().tag==q('p'):
                nxt=tail.getnext();parent.remove(tail);tail=nxt
            if tail is not None and tail.tag==q('p'):
                if blank_spacer(tail) and (tail.getnext() is None or tail.getnext().tag==q('sectPr')):
                    # No following content means there is no outside gap to
                    # show. Retain Word's small required final paragraph so a
                    # full landscape table cannot create an empty extra page.
                    pp=child(tail,'pPr');child(pp,'pStyle',val='VirkeMellomrom')
                    child(pp,'spacing',before=0,after=0,line=20,lineRule='exact')
                    child(child(pp,'rPr'),'sz',val=2)
                elif not txt(tail):gap_paragraph(tail,kind)
                else:child(child(tail,'pPr'),'spacing',before=360 if kind=='table' else 180)
    order(root)

def toc_gallery(glossary):
    source=block(glossary,'Innholdsfortegnelse')
    extra=deepcopy(source);pr=extra.find(q('docPartPr'))
    child(pr,'name',val='Virke innholdsfortegnelse')
    child(pr,'description',val='Virkes innholdsfortegnelse med tre overskriftsnivåer og levende sidetall.')
    child(child(pr,'category'),'gallery',val='tblOfContents')
    # A gallery entry is a normal building block; autoExp/toolbar belong to
    # the original AutoText entry. The Mac gallery still only lists built-ins.
    types=child(pr,'types');types.clear();child(types,'type',val='normal')
    child(pr,'guid',val='{'+str(uuid.uuid5(uuid.NAMESPACE_URL,'virke-toc-gallery-2026-10-27')).upper()+'}')
    glossary.find(q('docParts')).append(extra)

def build_one(source, dest):
    with ZipFile(source) as z: original={n:z.read(n) for n in z.namelist()}
    data=dict(original)
    for name in ['word/styles.xml','word/glossary/styles.xml']:
        r=E.fromstring(data[name]);revise_styles(r);data[name]=xml(r)
    for name in ['word/document.xml','word/glossary/document.xml']:
        r=E.fromstring(data[name]);revise_body(r);revise_gaps(r,glossary='glossary/' in name)
        if name.endswith('glossary/document.xml'):toc_gallery(r)
        data[name]=xml(r)
    for name in ['word/numbering.xml','word/glossary/numbering.xml']:
        if name in data:
            r=E.fromstring(data[name]);revise_numbering(r);data[name]=xml(r)
    if 'customUI/customUI.xml' in data:data['customUI/customUI.xml']=RIBBON.encode()
    if source.suffix=='.docx':
        settings=E.fromstring(data['word/settings.xml'])
        remove(settings,'attachedTemplate','linkStyles')
        child(settings,'updateFields',val='false')
        data['word/settings.xml']=xml(settings)
        relname='word/_rels/settings.xml.rels'
        if relname in data:
            rels=E.fromstring(data[relname])
            for rel in list(rels):
                if rel.get('Type','').endswith('/attachedTemplate'):rels.remove(rel)
            data[relname]=xml(rels)
    for n,raw in original.items():
        if any(x in n for x in ['/charts/','/embeddings/','/media/','/theme/']):assert data[n]==raw,n
    for n,raw in data.items():
        if n.endswith(('.xml','.rels')):E.fromstring(raw)
    dest.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(dest,'w',ZIP_DEFLATED) as z:
        for n,raw in data.items():z.writestr(n,raw)
    return {'file':str(dest.relative_to(ROOT)), 'sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),
            'changed_parts':[n for n in data if data[n]!=original.get(n)]}

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--master-only',action='store_true');args=parser.parse_args()
    files=[SOURCE/'Virke_master_test.dotx'] if args.master_only else sorted(p for p in SOURCE.iterdir() if p.suffix in ['.docx','.dotx'] and 'brukerveiledning' not in p.name)
    result=[build_one(p,DEST/p.name) for p in files]
    (QA/'candidate-build.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
    print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
