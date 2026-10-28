"""Surgical OOXML revision of the preserved 16 October delivery.

Keeps original parts, numbering and drawing relationships. Never round-trips
through python-docx; Word itself is used for interactive regression checks.
"""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
import argparse, hashlib, json, uuid, posixpath

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / 'Virke-alle-Figma-eksempler-2026-10-16'
DEST = ROOT / 'Virke-revidert-2026-10-21'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'r': R}
q = lambda s: '{'+W+'}'+s

def xml(r): return E.tostring(r, encoding='UTF-8', xml_declaration=True, standalone=True)
def txt(r): return ''.join(r.xpath('.//w:t/text()', namespaces=NS))
def child(r, tag, **attrs):
    n=r.find(q(tag))
    if n is None:
        n=E.SubElement(r,q(tag))
        if tag in ('pPr','rPr','tcPr','trPr','tblPr'):r.remove(n);r.insert(0,n)
    for k,v in attrs.items():n.set(q(k),str(v))
    return n

def remove(r,*tags):
    for tag in tags:
        for e in r.findall(q(tag)):r.remove(e)

STYLE_ORDER='name aliases basedOn next link autoRedefine hidden uiPriority semiHidden unhideWhenUsed qFormat locked personal personalCompose personalReply rsid pPr rPr tblPr trPr tcPr tblStylePr'.split()
P_ORDER='pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split()
R_ORDER='rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath'.split()
TC_ORDER='cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark'.split()

def order(root):
    for tag,seq in [('style',STYLE_ORDER),('pPr',P_ORDER),('rPr',R_ORDER),('tcPr',TC_ORDER)]:
        for n in root.iter(q(tag)):
            n[:]=sorted(n,key=lambda e:seq.index(E.QName(e).localname) if E.QName(e).localname in seq else len(seq))

QUICK=['Normal','Overskrift1','Overskrift2','Overskrift3','Listeavsnitt','VirkeListe','VirkeHjelpetekst','Bildetekst']

def revise_styles(root):
    latent=root.find(q('latentStyles'))
    if latent is not None:
        latent.set(q('defQFormat'),'0');latent.set(q('defSemiHidden'),'1');latent.set(q('defUnhideWhenUsed'),'0')
        for exc in latent:
            exc.set(q('qFormat'),'0');exc.set(q('semiHidden'),'1');exc.set(q('unhideWhenUsed'),'0')
    for s in root.findall(q('style')):
        sid=s.get(q('styleId'));remove(s,'qFormat','unhideWhenUsed')
        if sid in QUICK:
            remove(s,'hidden');child(s,'semiHidden',val=0);child(s,'uiPriority',val=QUICK.index(sid)+1);child(s,'qFormat')
        else:
            child(s,'semiHidden');child(s,'uiPriority',val=99)
        if sid in ('Overskrift1','Overskrift2','Tittel','Overskrift1Tegn','Overskrift2Tegn','TittelTegn'):
            child(child(s,'rPr'),'spacing',val=0)
        if sid=='Listeavsnitt':child(s,'next',val='Listeavsnitt')
        if sid in ('Virke1','Virke2'):
            child(child(s,'tcPr'),'vAlign',val='top')
            mar=child(child(s,'tblPr'),'tblCellMar')
            child(mar,'top',w=150,type='dxa');child(mar,'bottom',w=90,type='dxa')
            for cond in s.findall(q('tblStylePr')):
                child(child(cond,'tcPr'),'vAlign',val='top')
                if cond.get(q('type'))=='firstRow':
                    # Bold is a toggle in the style cascade. Applying it in both
                    # the conditional table style and paragraph style cancelled
                    # it after Word replaced a placeholder with typed text.
                    # The header paragraph style owns bold for every header cell.
                    remove(child(cond,'rPr'),'b','bCs')
        if sid in ('INNH1','INNH2','INNH3'):
            rp=child(s,'rPr');child(rp,'caps',val=0);child(rp,'smallCaps',val=0)
            child(rp,'sz',val=15);child(rp,'color',val='000000' if sid=='INNH1' else '616161')
            pp=child(s,'pPr');child(pp,'spacing',line=240,lineRule='atLeast',before=150 if sid=='INNH1' else 0,after=90)
            if sid!='INNH1':child(pp,'ind',left=600 if sid=='INNH2' else 960,hanging=180)
        if sid=='VirkeTabelloverskrift':
            child(child(s,'rPr'),'b',val=1);child(child(s,'rPr'),'bCs',val=1)
    order(root)


def revise_settings(root):
    # Word for Mac's Recommended filter omits this element. Explicitly mixing
    # visibleStyles/top3HeadingStyles hid Normal and other chosen styles.
    remove(root,'stylePaneFormatFilter')
    child(root,'stylePaneSortMethod',val='0001')
    remove(root,'linkStyles')


def set_bold(rp):
    child(rp,'b',val=1);child(rp,'bCs',val=1)
    f=child(rp,'rFonts')
    if not any('Arial' in v for v in f.attrib.values()):
        for k in ['ascii','hAnsi','cs','eastAsia']:f.attrib.pop(q(k),None)
        child(rp,'rFonts',asciiTheme='majorHAnsi',hAnsiTheme='majorHAnsi',cstheme='majorBidi')


def revise_tables(root):
    for t in root.iter(q('tbl')):
        pr=child(t,'tblPr');sid=pr.find(q('tblStyle'))
        data=sid is not None and sid.get(q('val')) in ('Virke1','Virke2')
        if data:
            mar=child(pr,'tblCellMar');child(mar,'top',w=150,type='dxa');child(mar,'bottom',w=90,type='dxa')
        for c in t.findall('.//'+q('tc')):
            cp=child(c,'tcPr');child(cp,'vAlign',val='top')
            if data:
                mar=child(cp,'tcMar');child(mar,'top',w=150,type='dxa');child(mar,'bottom',w=90,type='dxa')
        rows=t.findall(q('tr'))
        if data and rows:
            first=rows[0];child(child(first,'trPr'),'tblHeader')
            for c in first.findall(q('tc')):
                for p in c.iter(q('p')):
                    pp=child(p,'pPr');child(pp,'pStyle',val='VirkeTabelloverskrift');set_bold(child(pp,'rPr'))
                    for r in p.iter(q('r')):set_bold(child(r,'rPr'))
                    for sp in p.iter(q('sdtPr')):set_bold(child(sp,'rPr'))


def block(root,name):
    return next(d for d in root.find(q('docParts')) if d.find('w:docPartPr/w:name',NS).get(q('val'))==name)


def add_placeholder(glossary,label):
    name='VIRKE_PH_'+uuid.uuid5(uuid.NAMESPACE_URL,'virke-2026:'+label).hex.upper()
    parts=glossary.find(q('docParts'))
    if not any(d.find('w:docPartPr/w:name',NS).get(q('val'))==name for d in parts):
        d=E.SubElement(parts,q('docPart'));p=E.SubElement(d,q('docPartPr'));child(p,'name',val=name)
        cat=child(p,'category');child(cat,'name',val='General');child(cat,'gallery',val='placeholder')
        child(child(p,'types'),'type',val='bbPlcHdr');child(p,'guid',val='{'+str(uuid.uuid5(uuid.NAMESPACE_URL,name)).upper()+'}')
        b=child(d,'docPartBody');p=E.SubElement(b,q('p'));r=E.SubElement(p,q('r'));child(child(r,'rPr'),'color',val='C9007F');E.SubElement(r,q('t')).text=label
    return name


def field(glossary,label):
    s=E.Element(q('sdt'));p=child(s,'sdtPr');child(child(p,'rPr'),'color',val='000000');child(p,'alias',val=label.strip('[]'))
    child(child(p,'placeholder'),'docPart',val=add_placeholder(glossary,label));child(p,'showingPlcHdr');child(p,'text')
    c=child(s,'sdtContent');r=E.SubElement(c,q('r'));child(child(r,'rPr'),'color',val='C9007F');E.SubElement(r,q('t')).text=label
    return s


def para(style,text=None,glossary=None):
    p=E.Element(q('p'));child(child(p,'pPr'),'pStyle',val=style)
    if glossary is not None:p.append(field(glossary,text))
    elif text is not None:E.SubElement(E.SubElement(p,q('r')),q('t')).text=text
    return p


def signatures(root,glossary):
    for p in list(root.iter(q('p'))):
        value=txt(p).strip()
        if value in ('Fornavn, Etternavn','Styreleder','Styremedlem') and p.getparent().tag==q('tc'):
            pp=child(p,'pPr');child(pp,'pStyle',val='VirkeSignatur')
            for n in list(p):
                if n is not pp:p.remove(n)
            p.append(field(glossary,'[Fornavn og etternavn]' if value=='Fornavn, Etternavn' else '[Stillingstittel]'))


def add_agreement(glossary):
    d=deepcopy(block(glossary,'Signatur'));pr=d.find(q('docPartPr'));child(pr,'name',val='Avtale');child(pr,'description',val='Redigerbart utgangspunkt med avtalepartene, overskrifter og signaturfelt.')
    child(pr,'guid',val='{'+str(uuid.uuid5(uuid.NAMESPACE_URL,'virke-avtale-2026')).upper()+'}')
    b=d.find(q('docPartBody'));signature=deepcopy(b.find(q('tbl')));b.clear()
    for sty,label in [('Overskrift1','[Avtaletittel]'),('Normal','Mellom'),('Normal','[Avtalepart 1]'),('Normal','og'),('Normal','[Avtalepart 2]'),('Overskrift2','1. Avtalens formål'),('Normal','[Beskriv formålet med avtalen]'),('Overskrift2','2. Avtalens innhold'),('Normal','[Skriv avtalens innhold]')]:
        p=para(sty,label,glossary if label.startswith('[') else None)
        if label in ('[Beskriv formålet med avtalen]','[Skriv avtalens innhold]'):
            # Remove the placeholder control on first input so Enter creates
            # ordinary paragraphs. A permanent single-line control blocks Enter.
            sp=p.find('w:sdt/w:sdtPr',NS)
            sp.insert(list(sp).index(sp.find(q('showingPlcHdr'))),E.Element(q('temporary')))
        b.append(p)
    b.append(signature);b.append(para('Normal'))
    glossary.find(q('docParts')).append(d)
    for old,new in [('Mergefield 1','Møteinformasjon – enkel'),('Mergefield 2','Møteinformasjon – detaljert'),('Innhold_Virke','Innholdsfortegnelse')]:
        child(block(glossary,old).find(q('docPartPr')),'name',val=new)
    # Keep the final numbered paragraph's paragraph mark inside the block.
    # Without a trailing Normal paragraph Word merges it with the destination
    # paragraph and drops level 3's style/numbering on AutoText insertion.
    block(glossary,'Lovtekst').find(q('docPartBody')).append(para('Normal'))


def fix_heading_tracking(root):
    for p in root.iter(q('p')):
        sid=p.find('w:pPr/w:pStyle',NS)
        if sid is not None and sid.get(q('val')) in ('Tittel','Overskrift1','Overskrift2'):
            for rp in p.iter(q('rPr')):child(rp,'spacing',val=0)


def refresh_ids(root):
    for i,p in enumerate(root.iter(q('sdtPr')),181000001):child(p,'id',val=i)

RIBBON='''<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon><tabs><tab id="VirkeTab" label="Virke" insertAfterMso="TabHome"><group id="VirkeInsert" label="Ferdig innhold"><button idMso="AutoManager" label="Virke-blokker" size="large" screentip="Sett inn en Virke-blokk" supertip="Velg Virke_master_test.dotx under Søk i. Velg blokk og trykk Sett inn. Plasser først markøren i et tomt avsnitt."/></group><group id="VirkeFields" label="Innholdsfortegnelse"><button idMso="FieldsUpdate" label="Oppdater felt" size="large" screentip="Oppdater valgt felt" supertip="Klikk i innholdsfortegnelsen først. Velg Oppdater hele tabellen når overskrifter er endret."/></group></tab></tabs></ribbon></customUI>'''


def validate(data,original):
    for n,raw in data.items():
        if n.endswith(('.xml','.rels')):
            r=E.fromstring(raw)
            if n.endswith('.rels'):
                scope=posixpath.dirname(posixpath.dirname(n))
                for rel in r:
                    if rel.get('TargetMode')=='External':continue
                    target=rel.get('Target','');resolved=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(scope,target))
                    assert resolved in data,(n,resolved)
    for n,raw in original.items():
        if any(x in n for x in ('/charts/','/embeddings/','/media/','/numbering.xml')):assert data[n]==raw,n
    g=E.fromstring(data['word/glossary/document.xml'])
    user=g.xpath('w:docParts/w:docPart[w:docPartPr/w:category/w:gallery[@w:val="autoTxt"]]',namespaces=NS)
    assert len(user)==21,len(user)
    names={d.find('w:docPartPr/w:name',NS).get(q('val')) for d in g.find(q('docParts'))}
    for part in ['word/document.xml','word/glossary/document.xml']:
        r=E.fromstring(data[part])
        for ref in r.xpath('//w:sdtPr/w:placeholder/w:docPart/@w:val',namespaces=NS):assert ref in names,ref
    assert b'QA-test' not in data['word/document.xml']


def reviewed_toc_cache(current, cache_path):
    """Restore Word's reviewed field result only for otherwise identical content."""
    cached = cache_path.read_bytes()
    def without_toc_and_bookmarks(raw):
        root = E.fromstring(raw)
        for sdt in list(root.findall('w:body/w:sdt', NS)):
            if any('TOC ' in (i.text or '') for i in sdt.iter(q('instrText'))):
                sdt.getparent().remove(sdt)
        for tag in ('bookmarkStart', 'bookmarkEnd'):
            for node in list(root.iter(q(tag))):
                node.getparent().remove(node)
        return E.tostring(root, method='c14n')
    assert without_toc_and_bookmarks(current) == without_toc_and_bookmarks(cached), 'Content changed: update and review TOC in Word again'
    return cached


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--native-toc-cache', type=Path, help='Reviewed Word document.xml for example 07; requires unchanged non-TOC content')
    args = parser.parse_args()
    report=[];DEST.mkdir(exist_ok=True)
    for f in sorted(SOURCE.iterdir()):
        if f.suffix not in ('.docx','.dotx'):continue
        with ZipFile(f) as z:original={n:z.read(n) for n in z.namelist()}
        data=dict(original);glossary=E.fromstring(data['word/glossary/document.xml']);doc=E.fromstring(data['word/document.xml'])
        signatures(doc,glossary);add_agreement(glossary)
        # The large title is deliberately a level-one heading; TOC includes it.
        if f.name.startswith('07_'):
            for s in doc.xpath('//w:body/w:p/w:pPr/w:pStyle[@w:val="Tittel"]',namespaces=NS):s.set(q('val'),'Overskrift1')
        for n,r in [('word/document.xml',doc),('word/glossary/document.xml',glossary)]:
            revise_tables(r);fix_heading_tracking(r);refresh_ids(r);order(r);data[n]=xml(r)
        for n in ('word/styles.xml','word/glossary/styles.xml'):
            r=E.fromstring(data[n]);revise_styles(r);data[n]=xml(r)
        for n in ('word/settings.xml','word/glossary/settings.xml'):
            if n in data:r=E.fromstring(data[n]);revise_settings(r);data[n]=xml(r)
        if f.suffix=='.dotx':
            r=E.fromstring(data['docProps/core.xml']);title=r.find('{http://purl.org/dc/elements/1.1/}title')
            if title is not None:title.text=None
            data['docProps/core.xml']=xml(r)
            data['customUI/customUI.xml']=RIBBON.encode()
            rel=E.fromstring(data['_rels/.rels']);E.SubElement(rel,'{'+PKG+'}Relationship',Id='VirkeRibbon',Type='http://schemas.microsoft.com/office/2006/relationships/ui/extensibility',Target='customUI/customUI.xml');data['_rels/.rels']=xml(rel)
        if f.name.startswith('07_') and args.native_toc_cache:
            data['word/document.xml'] = reviewed_toc_cache(data['word/document.xml'], args.native_toc_cache)
        validate(data,original)
        out=DEST/f.name
        with ZipFile(out,'w',ZIP_DEFLATED) as z:
            for n,raw in data.items():z.writestr(n,raw)
        report.append({'file':f.name,'sha256':hashlib.sha256(out.read_bytes()).hexdigest(),'changed_parts':[n for n in data if data[n]!=original.get(n)],'structural_validation':'passed','word_qa':'pending'})
    (ROOT/'Word-QA-2026-10-21/revisjon.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(f'Revised {len(report)} Word files. Native Word verification pending.')

if __name__=='__main__':main()
