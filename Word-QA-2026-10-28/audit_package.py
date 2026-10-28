from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import json, hashlib

ROOT=Path(__file__).resolve().parents[1]
QA=ROOT/'Word-QA-2026-10-28'
DEST=ROOT/'Virke-revidert-2026-10-28'
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
q=lambda x:'{'+NS['w']+'}'+x
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
results=[]
for p in sorted(DEST.iterdir()):
    if p.suffix not in ['.docx','.dotx'] or 'brukerveiledning' in p.name:continue
    old=ROOT/'Virke-revidert-2026-10-27'/p.name
    with ZipFile(p) as z,ZipFile(old) as prior:
        for name in z.namelist():
            if name.endswith(('.xml','.rels')):E.fromstring(z.read(name))
            if any(key in name for key in ['/charts/','/media/','/embeddings/','/theme/']):
                assert z.read(name)==prior.read(name),(p.name,name)
        for part in ['word/styles.xml','word/glossary/styles.xml']:
            r=E.fromstring(z.read(part))
            ids=[s.get(q('styleId')) for s in r.findall(q('style'))];assert len(ids)==len(set(ids))
            normal=r.find('w:style[@w:styleId="Normal"]',NS)
            assert normal.find(q('rPr')) is None and normal.find(q('pPr')) is None
            assert r.find('w:docDefaults/w:rPrDefault/w:rPr/w:sz',NS).get(q('val'))=='18'
            s=r.find('w:style[@w:styleId="Virke1"]',NS)
            assert s.find('w:rPr/w:sz',NS).get(q('val'))=='15'
            assert s.find('w:tblStylePr[@w:type="firstRow"]/w:rPr/w:b',NS).get(q('val'))=='1'
            for side,value in [('top','105'),('bottom','165'),('left','120'),('right','120')]:
                assert s.find('w:tblPr/w:tblCellMar/w:'+side,NS).get(q('w'))==value
            for sid in ['VirkeTabelltekst','VirkeTabelloverskrift']:
                assert r.find('w:style[@w:styleId="'+sid+'"]/w:rPr/w:b',NS) is None
        settings=E.fromstring(z.read('word/settings.xml'))
        assert settings.find('w:defaultTableStyle',NS).get(q('val'))=='Virke1'
        if p.suffix=='.docx':
            assert settings.find(q('attachedTemplate')) is None
            assert settings.find(q('linkStyles')) is None
            assert settings.find(q('updateFields')).get(q('val')) in ['false','0']
            for part in z.namelist():
                if part.endswith('.rels'):
                    assert not any(x.get('TargetMode')=='External' for x in E.fromstring(z.read(part)))
        g=E.fromstring(z.read('word/glossary/document.xml'))
        blocks=[x for x in g.findall('w:docParts/w:docPart',NS) if x.find('w:docPartPr/w:category/w:gallery',NS).get(q('val'))=='autoTxt']
        assert len(blocks)==21
        d=E.fromstring(z.read('word/document.xml'))
        for part in ['word/document.xml','word/glossary/document.xml']:
            r=E.fromstring(z.read(part))
            for table in r.findall('.//w:tbl',NS):
                st=table.find('w:tblPr/w:tblStyle',NS)
                if st is None or st.get(q('val'))!='Virke1':continue
                header=table.find(q('tr'))
                assert not header.findall('.//w:rPr/w:b',NS)
                assert not table.findall('w:tr/w:trPr/w:trHeight',NS)
                for margin in table.findall('w:tblPr/w:tblCellMar',NS)+table.findall('w:tr/w:tc/w:tcPr/w:tcMar',NS):
                    assert margin.find(q('top')).get(q('w'))=='105'
                    assert margin.find(q('bottom')).get(q('w'))=='165'
        controls=len(d.findall('.//w:sdt',NS))
        olddoc=E.fromstring(prior.read('word/document.xml'))
        delta=controls-len(olddoc.findall('.//w:sdt',NS))
        assert delta==(-9 if p.name.startswith('04_') else 0),(p.name,delta)
        results.append({'file':p.name,'sha256':sha(p),'public_components':21,'content_controls':controls,'control_delta':delta,'style_and_package_audit':'passed'})
assert len(results)==18
guide=ROOT/'output/docx/Virke_mastermal_brukerveiledning.docx'
with ZipFile(guide) as z:
    d=E.fromstring(z.read('word/document.xml'))
    names=d.xpath('//w:bookmarkStart/@w:name',namespaces=NS)
    assert len({x for x in names if x.startswith('GuideModule')})==21
    for name in z.namelist():
        if name.endswith(('.xml','.rels')):E.fromstring(z.read(name))
        if name.endswith('.rels'):assert not any(x.get('TargetMode')=='External' for x in E.fromstring(z.read(name)))
record={'date':'2026-10-28','status':'passed_structural_audit','documents':results,'guide_components':21,'guide_sha256':sha(guide),'note':'Structural assertions supplement native Word and visual checks; they do not certify unperformed UI tests.'}
(QA/'package-audit.json').write_text(json.dumps(record,ensure_ascii=False,indent=2)+'\n')
print('18 master/example packages and 21 guide components: structural audit passed.')
