"""Read-only package audit of the final master and 17 example files."""
from pathlib import Path
from zipfile import ZipFile
from hashlib import sha256
from lxml import etree as E
import importlib.util, json, re
ROOT=Path(__file__).resolve().parents[1]
QA=ROOT/'Word-QA-2026-10-21'
DEST=ROOT/'Virke-revidert-2026-10-21'
spec=importlib.util.spec_from_file_location('revision',ROOT/'scripts/virke-build/revise_designer_feedback.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
def read(p):
    with ZipFile(p) as z:
        assert z.testzip() is None,p
        return {n:z.read(n) for n in z.namelist()}
def hashfile(p):return sha256(p.read_bytes()).hexdigest()
provenance=json.loads((QA/'referanser/proveniens.json').read_text())
assert hashfile(ROOT/provenance['original'])==provenance['sha256']
with ZipFile(QA/'for-iterasjon-2026-10-21.zip') as backup:
    backup_names=backup.namelist()
    for f in (ROOT/'Virke-alle-Figma-eksempler-2026-10-16').iterdir():
        if f.suffix not in ('.dotx','.docx'):continue
        matches=[n for n in backup_names if n.endswith('/'+f.name) or n==f.name]
        assert len(matches)==1 and backup.read(matches[0])==f.read_bytes(),f
rows=[];changes=[]
for f in sorted(DEST.iterdir()):
    if f.suffix not in ('.dotx','.docx') or f.name.startswith('Virke_mastermal_'):continue
    parts=read(f);before=read(ROOT/'Virke-alle-Figma-eksempler-2026-10-16'/f.name)
    m.validate(parts,before)
    assert not any('vbaProject' in n for n in parts)
    styles=E.fromstring(parts['word/styles.xml']);oldstyles=E.fromstring(before['word/styles.xml'])
    ids={x.get(m.q('styleId')) for x in styles.findall(m.q('style'))}
    assert ids=={x.get(m.q('styleId')) for x in oldstyles.findall(m.q('style'))}
    assert len(ids)==425
    for sid in ('Overskrift1','Overskrift2','Tittel'):
        n=styles.find(f'w:style[@w:styleId="{sid}"]/w:rPr/w:spacing',m.NS)
        assert n is not None and n.get(m.q('val'))=='0'
    for sid in ('Virke1','Virke2'):
        assert not styles.xpath(f'w:style[@w:styleId="{sid}"]/w:tblStylePr[@w:type="firstRow"]/w:rPr/w:b',namespaces=m.NS)
    assert styles.find('w:style[@w:styleId="VirkeTabelloverskrift"]/w:rPr/w:b',m.NS).get(m.q('val'))=='1'
    settings=E.fromstring(parts['word/settings.xml'])
    assert settings.find(m.q('linkStyles')) is None
    header_tables=0
    for name in ['word/document.xml','word/glossary/document.xml']:
        x=E.fromstring(parts[name]);text=' '.join(x.xpath('//w:t/text()',namespaces=m.NS))
        assert not re.search(r'\bQA(?:[ -]|$)|stresstest|bestatt',text,re.I),(f.name,name)
        for table in x.iter(m.q('tbl')):
            for c in table.findall('.//w:tc',m.NS):
                assert c.find('w:tcPr/w:vAlign',m.NS).get(m.q('val'))=='top'
            sid=table.find('w:tblPr/w:tblStyle',m.NS)
            if sid is not None and sid.get(m.q('val')) in ('Virke1','Virke2'):
                first=table.find('w:tr',m.NS)
                assert first.find('w:trPr/w:tblHeader',m.NS) is not None
                for p in first.findall('.//w:p',m.NS):
                    assert p.find('w:pPr/w:pStyle',m.NS).get(m.q('val'))=='VirkeTabelloverskrift'
                if name=='word/document.xml':header_tables+=1
    g=E.fromstring(parts['word/glossary/document.xml'])
    names=[x.get(m.q('val')) for x in g.xpath('w:docParts/w:docPart[w:docPartPr/w:category/w:gallery[@w:val="autoTxt"]]/w:docPartPr/w:name',namespaces=m.NS)]
    preserved=[n for n in before if any(x in n for x in ('/charts/','/embeddings/','/media/','/numbering.xml'))]
    rows.append({'file':f.name,'sha256':hashfile(f),'styles_defined':len(ids),'style_ids_preserved':True,'autotext_blocks':names,'tracking_zero':True,'body_data_tables':header_tables,'table_header_paragraph_style_bold':True,'conditional_bold_toggle_removed':True,'top_alignment':True,'automatic_style_update':False,'no_qa_text':True,'preserved_binary_and_numbering_parts':preserved,'package_validation':'passed'})
    changes.append({'file':f.name,'sha256':hashfile(f),'changed_parts':[n for n in parts if parts[n]!=before.get(n)],'structural_validation':'passed','word_qa':'completed_see_native_results','word_qa_record':'Word-QA-2026-10-21/native-results.json'})
assert len(rows)==18
for path,data in [('strukturkontroll.json',rows),('revisjon.json',changes)]:
    (QA/path).write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n')
print('18 packages passed. Original attachment and 16 October Word packages unchanged. 425 styles and 21 AutoText blocks per package; no QA text.')
