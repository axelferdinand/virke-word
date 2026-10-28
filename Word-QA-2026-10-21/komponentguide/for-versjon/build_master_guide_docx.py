"""Build the editable guide by cloning the tested Virke master package.

Only documented content slots and new list/link instances change. The master,
its styles, glossary, charts, embeddings and existing numbering stay intact.
Export the resulting document from Word; build_master_guide.py records that PDF.
"""
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import json
from lxml import etree as E

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'output/docx/Virke_mastermal_brukerveiledning.docx'
TMP = ROOT / 'tmp/master-guide-docx'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'w': W}
def q(tag): return f'{{{W}}}{tag}'
def el(tag, **attrs):
    return E.Element(q(tag), {q(k): str(v) for k, v in attrs.items()})
def xml(root): return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)

def runs(parent, text):
    """Only bold and line breaks from the local JSON content are supported."""
    fragment = E.fromstring(('<root>'+text.replace('&', '&amp;')+'</root>').encode())
    def emit(value, bold=False):
        if not value: return
        chunks = value.split('\n')
        for idx, chunk in enumerate(chunks):
            r = el('r')
            if bold:
                rp = el('rPr'); rp.append(el('b', val=1)); r.append(rp)
            if idx: r.append(el('br'))
            t = el('t'); t.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); t.text=chunk
            r.append(t); parent.append(r)
    emit(fragment.text)
    for child in fragment:
        if child.tag != 'b': raise ValueError(f'Unsupported inline tag: {child.tag}')
        emit(child.text, True); emit(child.tail)

def para(text='', style='Normal', page_break=False, num_id=None):
    p=el('p'); pr=el('pPr'); pr.append(el('pStyle', val=style)); p.append(pr)
    if page_break: pr.append(el('pageBreakBefore'))
    if num_id is not None:
        np=el('numPr'); np.extend([el('ilvl', val=0),el('numId',val=num_id)]); pr.append(np)
    runs(p,text)
    return p

def fill_cell(tc, text, style):
    # A newly cloned table row contains static guide content, not form fields.
    # Its exact table/cell properties and source paragraph style remain in use.
    for child in list(tc):
        if child.tag != q('tcPr'): tc.remove(child)
    tc.append(para(text,style))

def main():
    content=json.loads((ROOT/'docs/brukerveiledning-master.json').read_text())
    master=ROOT/content['master']
    assert master.read_bytes()==(TMP/'reference.dotx').read_bytes(), 'Redistill changed master first'
    with ZipFile(master) as z: original={n:z.read(n) for n in z.namelist()}
    parts=original.copy()
    doc=E.fromstring(parts['word/document.xml']); body=doc.find('w:body',NS)
    sect=deepcopy(body.find('w:sectPr',NS)); body.clear()
    glossary=E.fromstring(parts['word/glossary/document.xml'])
    blocks={dp.find('w:docPartPr/w:name',NS).get(q('val')):dp.find('w:docPartBody',NS)
            for dp in glossary.findall('w:docParts/w:docPart',NS)}
    table_prototype=blocks['Tabell med 2 kolonner'].find('w:tbl',NS)
    note_prototype=blocks['Hjelpetekst'].find('w:tbl',NS)
    numbering=E.fromstring(parts['word/numbering.xml'])
    rels=E.fromstring(parts['word/_rels/document.xml.rels'])
    next_num=max(int(n.get(q('numId'))) for n in numbering.findall('w:num',NS))+1
    link_count=0

    for page_i,page in enumerate(content['pages']):
        body.append(para(page['title'],'Overskrift1',page_break=page_i>0))
        body.append(para(page['intro']))
        for block in page['blocks']:
            kind=block['type']
            if kind in ('heading','paragraph','small'):
                style={'heading':'Overskrift3','paragraph':'Normal','small':'Bildetekst'}[kind]
                body.append(para(block['text'],style))
            elif kind in ('bullets','steps'):
                num_id=21
                if kind=='steps':
                    num_id=next_num; next_num+=1
                    num=deepcopy(numbering.find('w:num[@w:numId="28"]',NS)); num.set(q('numId'),str(num_id))
                    for old in num.findall('w:lvlOverride',NS):num.remove(old)
                    override=el('lvlOverride',ilvl=0);override.append(el('startOverride',val=1));num.append(override)
                    numbering.append(num)
                for item in block['items']:body.append(para(item,'Listeavsnitt',num_id=num_id))
            elif kind=='table':
                table=deepcopy(table_prototype)
                head,row=deepcopy(table.findall('w:tr',NS)[0]),deepcopy(table.findall('w:tr',NS)[1])
                for old in table.findall('w:tr',NS):table.remove(old)
                widths=[round(9330*f) for f in block['widths']]; widths[-1]=9330-sum(widths[:-1])
                for gc,width in zip(table.findall('w:tblGrid/w:gridCol',NS),widths):gc.set(q('w'),str(width))
                for idx,values in enumerate([block['headers']]+block['rows']):
                    tr=deepcopy(head if idx==0 else row)
                    for tc,text,width in zip(tr.findall('w:tc',NS),values,widths):
                        tc.find('w:tcPr/w:tcW',NS).set(q('w'),str(width))
                        fill_cell(tc,text,'VirkeTabelloverskrift' if idx==0 else 'VirkeTabelltekst')
                    table.append(tr)
                body.append(table);body.append(para('','VirkeMellomrom'))
            elif kind=='note':
                table=deepcopy(note_prototype);tc=table.find('w:tr/w:tc',NS)
                fill_cell(tc,'<b>'+block['title']+'</b>\n'+block['text'],'VirkeHjelpetekst')
                body.append(table);body.append(para('','VirkeMellomrom'))
            elif kind=='sources':
                body.append(para('Mer hjelp fra Microsoft','Overskrift3'))
                for label,url in block['items']:
                    link_count+=1;rid=f'rIdGuideLink{link_count}'
                    rel=E.SubElement(rels,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
                    rel.set('Id',rid);rel.set('Type',R+'/hyperlink');rel.set('Target',url);rel.set('TargetMode','External')
                    p=para('','Bildetekst');link=el('hyperlink');link.set('{'+R+'}id',rid)
                    runs(link,label);p.append(link);body.append(p)
            else: raise ValueError(kind)
    body.append(sect)
    parts['word/document.xml']=xml(doc)
    parts['word/numbering.xml']=xml(numbering)
    parts['word/_rels/document.xml.rels']=xml(rels)
    for part,text in [('word/header1.xml','Virke Word · Brukerveiledning '+content['version']),
                      ('word/footer1.xml','Virke mastermal · '+content['version'])]:
        root=E.fromstring(parts[part])
        for sdt in root.findall('.//w:sdt',NS):
            ts=sdt.findall('w:sdtContent//w:t',NS)
            assert len(ts)==1, 'Expected verified plain text slot'
            ts[0].text=text
            for ph in sdt.findall('w:sdtPr/w:showingPlcHdr',NS):ph.getparent().remove(ph)
            for color in sdt.findall('w:sdtContent//w:color',NS):color.set(q('val'),'616161')
        parts[part]=xml(root)
    ct=parts['[Content_Types].xml'].replace(b'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',b'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
    parts['[Content_Types].xml']=ct
    core=E.fromstring(parts['docProps/core.xml'])
    for key,value in [('title',content['title']),('creator','Virke')]:
        node=core.find('{http://purl.org/dc/elements/1.1/}'+key)
        if node is None:node=E.SubElement(core,'{http://purl.org/dc/elements/1.1/}'+key)
        node.text=value
    parts['docProps/core.xml']=xml(core)
    OUT.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(OUT,'w',ZIP_DEFLATED) as z:
        for n,b in parts.items():z.writestr(n,b)
    allowed={'word/document.xml','word/header1.xml','word/footer1.xml','word/numbering.xml',
             'word/_rels/document.xml.rels','docProps/core.xml','[Content_Types].xml'}
    changed={n for n,b in parts.items() if b!=original[n]}
    assert changed<=allowed
    before=E.fromstring(original['word/numbering.xml'])
    assert all(E.tostring(a)==E.tostring(b) for a,b in zip(before,numbering)), 'Existing numbering changed'
    report={'master_sha256':sha256(master.read_bytes()).hexdigest(),'docx_sha256':sha256(OUT.read_bytes()).hexdigest(),
            'changed_parts':sorted(changed),'preserved_parts':len(original)-len(changed),
            'existing_numbering_preserved':True,'style_definitions':425,'logical_pages':len(content['pages']),
            'native_word_review':'pending','pdf_export':'pending'}
    (TMP/'build-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(OUT)

if __name__=='__main__':main()
