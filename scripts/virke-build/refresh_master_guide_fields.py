"""Import a reviewed native Word TOC cache without normalizing master parts."""
import argparse
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import json
from lxml import etree as E
from build_master_guide_docx import ROOT, OUT, TMP, NS, q, el, xml

def refresh(native):
    content=json.loads((ROOT/'docs/brukerveiledning-master.json').read_text())
    pages=len(content['pages'])
    footer_page=next(i for i,p in enumerate(content['pages'],1)
                     if any(b.get('name')=='Virke bunntekst' for b in p['blocks']))
    with ZipFile(OUT) as z: parts={n:z.read(n) for n in z.namelist()}
    with ZipFile(native) as z: native_doc=E.fromstring(z.read('word/document.xml'))
    doc=E.fromstring(parts['word/document.xml'])
    query='//w:sdt[w:sdtPr/w:docPartObj]'
    old=doc.xpath(query,namespaces=NS)[0]
    toc=deepcopy(native_doc.xpath(query,namespaces=NS)[0])
    assert 'GuideStyleSample' in ''.join(toc.xpath('.//w:instrText/text()',namespaces=NS))
    anchors=toc.xpath('.//w:hyperlink/@w:anchor',namespaces=NS)
    assert len(anchors)==3
    original_marks=[deepcopy(e) for e in old.xpath('.//w:bookmarkStart|.//w:bookmarkEnd',namespaces=NS)]
    assert len(original_marks)==2
    for e in toc.xpath('.//w:bookmarkStart|.//w:bookmarkEnd',namespaces=NS):e.getparent().remove(e)
    ps=toc.findall('w:sdtContent/w:p',NS)
    ps[0].insert(1,original_marks[0]);ps[-1].append(original_marks[1])
    toc.find('w:sdtPr/w:id',NS).set(q('val'),old.find('w:sdtPr/w:id',NS).get(q('val')))
    old.getparent().replace(old,toc)
    for i,name in enumerate(anchors,9001):
        source=native_doc.xpath('//w:bookmarkStart[@w:name=$name]',namespaces=NS,name=name)[0].getparent()
        text=''.join(source.xpath('.//w:t/text()',namespaces=NS))
        candidates=[p for p in doc.findall('w:body/w:p',NS) if ''.join(p.xpath('.//w:t/text()',namespaces=NS))==text]
        assert len(candidates)==1
        p=candidates[0];p.insert(1,el('bookmarkStart',id=i,name=name));p.append(el('bookmarkEnd',id=i))
    def cache(root,values):
        for instr in root.findall('.//w:instrText',NS):
            key=(instr.text or '').strip().upper()
            if key not in values:continue
            for sibling in instr.getparent().itersiblings():
                t=sibling.find('w:t',NS)
                if t is not None:t.text=str(values[key]);break
    cache(doc,{'PAGE':footer_page,'NUMPAGES':pages})
    parts['word/document.xml']=xml(doc)
    for part in ['word/footer1.xml','word/footerGuideLandscape.xml']:
        root=E.fromstring(parts[part]);cache(root,{'NUMPAGES':pages});parts[part]=xml(root)
    names=doc.xpath('//w:bookmarkStart/@w:name',namespaces=NS)
    assert len(names)==len(set(names)) and all(a in names for a in anchors)
    with ZipFile(OUT,'w',ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)
    report_path=TMP/'build-report.json';report=json.loads(report_path.read_text())
    report.update(docx_sha256=sha256(OUT.read_bytes()).hexdigest(),native_field_source=str(native),native_toc_entries=3)
    report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(OUT)

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--native-docx',type=Path,required=True)
    refresh(parser.parse_args().native_docx)
