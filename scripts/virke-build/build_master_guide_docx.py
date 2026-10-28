"""Build the editable Virke component guide from the retained master package.

The actual glossary blocks are cloned into the document. Source styles, themes,
charts, workbooks and glossary remain byte-identical. Only documented text slots,
new navigation/list instances and guide page furniture change.
"""
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import json
from lxml import etree as E

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT/'output/docx/Virke_mastermal_brukerveiledning.docx'
TMP = ROOT/'tmp/master-guide-docx'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'r': R, 'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
def q(tag): return f'{{{W}}}{tag}'
def el(tag, **attrs): return E.Element(q(tag), {q(k):str(v) for k,v in attrs.items()})
def xml(root): return E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
def runs(parent,text):
    fragment=E.fromstring(('<root>'+text.replace('&','&amp;')+'</root>').encode())
    def emit(value,bold=False):
        if not value:return
        for idx,chunk in enumerate(value.split('\n')):
            r=el('r')
            if bold:
                rp=el('rPr');rp.append(el('b',val=1));r.append(rp)
            if idx:r.append(el('br'))
            t=el('t');t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=chunk
            r.append(t);parent.append(r)
    emit(fragment.text)
    for child in fragment:
        if child.tag!='b':raise ValueError(child.tag)
        emit(child.text,True);emit(child.tail)
def para(text='',style='Normal',page_break=False,num_id=None):
    p=el('p');pr=el('pPr');pr.append(el('pStyle',val=style));p.append(pr)
    if page_break:pr.append(el('pageBreakBefore'))
    if num_id is not None:
        np=el('numPr');np.extend([el('ilvl',val=0),el('numId',val=num_id)]);pr.append(np)
    runs(p,text);return p

def filled_slot(node,text):
    """Change verified text slots in place, retaining control and paragraph XML."""
    ts=node.findall('.//w:t',NS)
    if not ts:raise ValueError('Missing text slot')
    ts[0].text=text
    for t in ts[1:]:t.text=''
    for sdt in node.iter(q('sdt')):
        for ph in sdt.findall('w:sdtPr/w:showingPlcHdr',NS):ph.getparent().remove(ph)
        for color in sdt.findall('w:sdtContent//w:color',NS):color.set(q('val'),'000000')

def fill_map(tree,mapping):
    queues={k:(list(v) if isinstance(v,list) else v) for k,v in mapping.items()}
    def replacement(key):
        value=queues[key]
        if isinstance(value,list):
            if not value:raise ValueError('Too few replacements for '+key)
            return value.pop(0)
        return value
    # Only matching plain text SDTs are filled. The TOC rich content is untouched.
    for sdt in tree.findall('.//w:sdt',NS):
        if sdt.findall('.//w:sdt',NS):continue
        text=''.join(sdt.xpath('w:sdtContent//w:t/text()',namespaces=NS))
        if text in queues:filled_slot(sdt,replacement(text))
    for t in tree.findall('.//w:t',NS):
        if any(a.tag==q('sdt') for a in t.iterancestors()):continue
        if t.text in queues:t.text=replacement(t.text)
    assert all(not isinstance(v,list) or not v for v in queues.values()), 'Unused replacement values'

def main():
    content=json.loads((ROOT/'docs/brukerveiledning-master.json').read_text())
    master=ROOT/content['master']
    assert master.read_bytes()==(TMP/'reference.dotx').read_bytes(), 'Master reference changed'
    with ZipFile(master) as z:original={n:z.read(n) for n in z.namelist()}
    parts=original.copy();doc=E.fromstring(parts['word/document.xml']);body=doc.find('w:body',NS)
    portrait=deepcopy(body.find('w:sectPr',NS));body.clear()
    landscape=deepcopy(portrait)
    landscape.find('w:pgSz',NS).attrib.update({q('w'):'16838',q('h'):'11906',q('orient'):'landscape'})
    landscape.find('w:pgMar',NS).attrib.update({q('left'):'709',q('right'):'709'})
    landscape.find('w:footerReference',NS).set('{'+R+'}id','rIdGuideLandscapeFooter')
    glossary=E.fromstring(parts['word/glossary/document.xml'])
    blocks={d.find('w:docPartPr/w:name',NS).get(q('val')):d.find('w:docPartBody',NS) for d in glossary.findall('w:docParts/w:docPart',NS)}
    public={d.find('w:docPartPr/w:name',NS).get(q('val')) for d in glossary.findall('w:docParts/w:docPart',NS)
            if not d.find('w:docPartPr/w:name',NS).get(q('val')).startswith('VIRKE_PH_')
            and d.find('w:docPartPr/w:category/w:gallery',NS).get(q('val'))=='autoTxt'}
    assert len(public)==21
    numbering=E.fromstring(parts['word/numbering.xml']);rels=E.fromstring(parts['word/_rels/document.xml.rels'])
    rel=E.SubElement(rels,'{'+PKG+'}Relationship',Id='rIdGuideLandscapeFooter',Type=R+'/footer',Target='footerGuideLandscape.xml')
    next_num=max(int(n.get(q('numId'))) for n in numbering.findall('w:num',NS))+1
    modules=[];locations={}
    for page_no,page in enumerate(content['pages'],1):
        for b in page['blocks']:
            if b['type']=='module' and b['name'] not in modules:
                modules.append(b['name']);locations[b['name']]=page_no
    assert set(modules)==public, 'Every public block must appear as an actual module'
    module_bookmarks={name:'GuideModule%02d'%(i+1) for i,name in enumerate(modules)}
    marked=set();coverage=[];bookmark_id=5000
    descriptions={
      'Hjelpetekst':'Kort instruks om utfylling','Innholdsfortegnelse':'Levende oversikt fra overskrifter','Lovtekst':'Nummererte punkter på tre nivåer',
      'Utfyllingsmal':'Ni navnefelt med faste etiketter','Møteinformasjon – enkel':'Møtedata og deltakere','Møteinformasjon – detaljert':'Dato, tid og sted på én linje',
      'Tabell':'Beløpsoppsett med sumlinje','Tabell med 2 kolonner':'To kolonner med overskriftsrad','Tabell med 3 kolonner':'Tre kolonner med overskriftsrad',
      'Tabell med 4 kolonner':'Fire kolonner med overskriftsrad','Tabell med 5 kolonner':'Fem kolonner med overskriftsrad','Store celler':'Kontrollpunkter og lengre merknader',
      'Tabell med 9 kolonner liggende':'Ulike kolonnebredder på liggende side','Tabell med 9 like kolonner liggende':'Like kolonnebredder på liggende side',
      'Saksliste':'Saker, beskrivelse og vedtak','Signatur':'Sted, dato og tre signaturfelt','Avtale':'Tittel, parter, avsnitt og signatur',
      'Diagram søyle':'Sammenligning av kategorier','Diagram sektor':'Fordeling av én total','Virke topptekst':'Tema eller dokumentkategori','Virke bunntekst':'Dokumentnavn og levende sidetall'}
    def mark_range(first,last,name):
        nonlocal bookmark_id
        bookmark_id+=1
        start=el('bookmarkStart',id=bookmark_id,name=name);end=el('bookmarkEnd',id=bookmark_id)
        first.insert(1 if first.find('w:pPr',NS) is not None else 0,start);last.append(end)
    def fresh_list():
        nonlocal next_num
        result=next_num;next_num+=1
        num=deepcopy(numbering.find('w:num[@w:numId="28"]',NS));num.set(q('numId'),str(result))
        for old in num.findall('w:lvlOverride',NS):num.remove(old)
        ov=el('lvlOverride',ilvl=0);ov.append(el('startOverride',val=1));num.append(ov);numbering.append(num)
        return result
    def make_rows(table,rows):
        originals=table.findall('w:tr',NS)
        assert len(originals)>=2
        for row in originals:table.remove(row)
        for i,values in enumerate(rows):
            tr=deepcopy(originals[0 if i==0 else 1]);cells=tr.findall('w:tc',NS)
            assert len(cells)==len(values)
            for cell,value in zip(cells,values):filled_slot(cell,value)
            table.append(tr)
    def add_module(b,page_no):
        name=b['name'];tree=deepcopy(blocks[name])
        if len(tree)>=2:
            before_tail=tree[-2].find('w:pPr/w:pStyle',NS)
            if (tree[-1].tag==q('p') and not tree[-1].xpath('.//w:t/text()',namespaces=NS)
                and before_tail is not None and before_tail.get(q('val')) in ['VirkeEtterTabell','VirkeEtterModul']):
                tree.remove(tree[-1])
        for tag in ['bookmarkStart','bookmarkEnd']:
            for node in tree.findall('.//w:'+tag,NS):node.getparent().remove(node)
        if b.get('fills'):fill_map(tree,b['fills'])
        table=tree.find('w:tbl',NS)
        if 'rows' in b:make_rows(table,b['rows'])
        if 'max_body_rows' in b:
            for tr in table.findall('w:tr',NS)[b['max_body_rows']+1:]:table.remove(tr)
        if b.get('toc_bookmark'):
            for ins in tree.findall('.//w:instrText',NS):
                if ins.text and ins.text.strip().startswith('TOC '):ins.text=' TOC \\o "1-3" \\h \\z \\u \\b '+b['toc_bookmark']+' '
            # Keep the outer field end in its own unbulleted paragraph. Word
            # otherwise leaves a blank TOC3 bullet after regenerating the cache.
            end=tree.findall('.//w:fldChar',NS)[-1]
            assert end.get(q('fldCharType'))=='end'
            end_run=end.getparent();end_run.getparent().remove(end_run)
            tail=para('','INNH1');tail.append(end_run)
            tree.find('w:sdt/w:sdtContent',NS).append(tail)
        ps=tree.findall('.//w:p',NS)
        if name not in marked:
            mark_range(ps[0],ps[-1],module_bookmarks[name]);marked.add(name)
        for child in list(tree):
            if child.tag!=q('sectPr'):body.append(child)
        # Body drawings already have identical targets in the master document rels.
        coverage.append({'module':name,'logical_page':page_no,'bookmark':module_bookmarks[name],
                         'source':'word/glossary/document.xml','actual_component':True})
    for idx,page in enumerate(content['pages']):
        orientation=page.get('orientation','portrait')
        prev=content['pages'][idx-1].get('orientation','portrait') if idx else orientation
        switched=idx and orientation!=prev
        if switched:
            p=para('','VirkeMellomrom');sect=deepcopy(landscape if prev=='landscape' else portrait)
            typ=el('type',val='nextPage');sect.insert(2,typ);p.find('w:pPr',NS).append(sect);body.append(p)
        title=para(page['title'],'Tittel',page_break=idx>0 and not switched)
        mark_range(title,title,'GuideChapter%02d'%(idx+1));body.append(title);body.append(para(page['intro']))
        for b in page['blocks']:
            kind=b['type']
            if kind in ['paragraph','small','heading']:
                style='Bildetekst' if kind=='small' else 'Normal'
                p=para('<b>'+b['text']+'</b>' if kind=='heading' else b['text'],style)
                if kind=='heading':
                    pr=p.find('w:pPr',NS);pr.append(el('keepNext'));pr.append(el('spacing',before=180,after=120))
                body.append(p)
            elif kind in ['bullets','steps']:
                nid=5 if kind=='bullets' else fresh_list()
                for item in b['items']:body.append(para(item,'Punktliste' if kind=='bullets' else 'Listeavsnitt',num_id=nid))
            elif kind=='module':add_module(b,idx+1)
            elif kind=='style_sample':
                ps=[para('Overskrift 1 viser hovedtemaet','Overskrift1'),para('Denne brødteksten bruker Normal. Bruk den til forklaringer og avsnitt. <b>Fet skrift</b> kan fremheve korte ord eller uttrykk.'),
                    para('Overskrift 2 deler inn temaet','Overskrift2'),para('Legg underkapitler på nivå 2. Da blir forholdet til hovedtemaet tydelig.'),
                    para('Overskrift 3 utdyper underkapitlet','Overskrift3'),para('Bruk tredje nivå når underkapitlet trenger en finere inndeling. Behold nivåene i rekkefølge.')]
                mark_range(ps[0],ps[-1],'GuideStyleSample');body.extend(ps)
            elif kind=='module_index':
                table=deepcopy(blocks['Tabell med 3 kolonner'].find('w:tbl',NS))
                rows=[['Byggeblokk','Bruksområde','Side']]+[[name,descriptions[name],str(locations[name])] for name in modules]
                make_rows(table,rows)
                widths=[4050,4480,800]
                for gc,width in zip(table.findall('w:tblGrid/w:gridCol',NS),widths):gc.set(q('w'),str(width))
                for row_no,row in enumerate(table.findall('w:tr',NS)):
                    for cell,width in zip(row.findall('w:tc',NS),widths):cell.find('w:tcPr/w:tcW',NS).set(q('w'),str(width))
                    if row_no:
                        cell=row.findall('w:tc',NS)[2];p=cell.find('w:p',NS)
                        for child in list(p):
                            if child.tag!=q('pPr'):p.remove(child)
                        field=el('fldSimple',instr=' PAGEREF '+module_bookmarks[modules[row_no-1]]+' \\h ')
                        runs(field,str(locations[modules[row_no-1]]));p.append(field)
                body.append(table);body.append(para('','VirkeEtterTabell'))
            else:raise ValueError(kind)
    body.append(deepcopy(landscape if content['pages'][-1].get('orientation')=='landscape' else portrait))
    # All top-level guide tables use the master's 9 pt preceding space.
    for table in body.findall('w:tbl',NS):
        prev=table.getprevious()
        if prev is None or prev.tag!=q('p'):continue
        pr=prev.find('w:pPr',NS)
        if pr is None:pr=el('pPr');prev.insert(0,pr)
        sp=pr.find('w:spacing',NS)
        if sp is None:sp=el('spacing');pr.append(sp)
        sp.set(q('after'),'180')
    # The master ends tables with an editable continuation paragraph. Here the
    # following guide paragraph is already supplied, so join them without an
    # extra empty line, retaining the module bookmark at the component's end.
    for p in list(body.findall('w:p',NS)):
        st=p.find('w:pPr/w:pStyle',NS)
        if st is None or st.get(q('val')) not in ['VirkeEtterTabell','VirkeEtterModul']:continue
        if p.xpath('.//w:t/text()',namespaces=NS):continue
        before=360 if st.get(q('val'))=='VirkeEtterTabell' else 180
        nxt=p.getnext();prev=p.getprevious()
        if nxt is None or nxt.tag!=q('p'):continue
        np=nxt.find('w:pPr',NS)
        if np is None:np=el('pPr');nxt.insert(0,np)
        sp=np.find('w:spacing',NS)
        if sp is None:sp=el('spacing');np.append(sp)
        sp.set(q('before'),str(max(before,int(sp.get(q('before'),'0')))))
        anchor=prev.findall('.//w:p',NS)[-1] if prev is not None and prev.tag==q('tbl') else prev
        for bm in list(p):
            if bm.tag in [q('bookmarkStart'),q('bookmarkEnd')]:anchor.append(bm)
        body.remove(p)
    assert marked==public
    # Each inserted control and drawing gets a distinct identity in the new document.
    identity=2000000000
    for node in doc.findall('.//w:sdtPr/w:id',NS):identity+=1;node.set(q('val'),str(identity))
    for i,node in enumerate(doc.findall('.//wp:docPr',NS),1):node.set('id',str(i))
    parts['word/document.xml']=xml(doc);parts['word/numbering.xml']=xml(numbering);parts['word/_rels/document.xml.rels']=xml(rels)
    for part,text in [('word/header1.xml','Virke Word · Komponentguide '+content['version']),('word/footer1.xml','Virke mastermal · '+content['version'])]:
        root=E.fromstring(parts[part])
        for sdt in root.findall('.//w:sdt',NS):
            ts=sdt.findall('w:sdtContent//w:t',NS);assert len(ts)==1
            filled_slot(sdt,text)
            for color in sdt.findall('w:sdtContent//w:color',NS):color.set(q('val'),'616161')
        for node in root.findall('.//w:sdtPr/w:id',NS):identity+=1;node.set(q('val'),str(identity))
        parts[part]=xml(root)
    footer=E.fromstring(parts['word/footer1.xml'])
    for tab in footer.findall('.//w:tabs/w:tab',NS):
        if tab.get(q('pos'))=='9330':tab.set(q('pos'),'15420')
    footer.find('.//w:tabs',NS).insert(0,el('tab',val='clear',pos='9330'))
    for node in footer.findall('.//w:sdtPr/w:id',NS):identity+=1;node.set(q('val'),str(identity))
    parts['word/footerGuideLandscape.xml']=xml(footer)
    ct=E.fromstring(parts['[Content_Types].xml'])
    for item in ct:
        if item.get('PartName')=='/word/document.xml':item.set('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
    E.SubElement(ct,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/word/footerGuideLandscape.xml',ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
    parts['[Content_Types].xml']=xml(ct)
    core=E.fromstring(parts['docProps/core.xml'])
    for key,value in [('title',content['title']),('creator','Virke')]:
        node=core.find('{http://purl.org/dc/elements/1.1/}'+key)
        if node is None:node=E.SubElement(core,'{http://purl.org/dc/elements/1.1/}'+key)
        node.text=value
    parts['docProps/core.xml']=xml(core)
    allowed={'word/document.xml','word/header1.xml','word/footer1.xml','word/footerGuideLandscape.xml','word/numbering.xml','word/_rels/document.xml.rels','docProps/core.xml','[Content_Types].xml'}
    changed={n for n,b in parts.items() if original.get(n)!=b};assert changed<=allowed
    before=E.fromstring(original['word/numbering.xml']);assert all(E.tostring(a)==E.tostring(b) for a,b in zip(before,numbering))
    # Document-level component relationships must resolve without external assets.
    relation_ids={rel.get('Id') for rel in rels}
    for element in doc.iter():
        for key,value in element.attrib.items():
            if key.startswith('{'+R+'}') and E.QName(key).localname in ['id','embed','link']:assert value in relation_ids,(key,value)
    OUT.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(OUT,'w',ZIP_DEFLATED) as z:
        for n,b in parts.items():z.writestr(n,b)
    report={'master_sha256':sha256(master.read_bytes()).hexdigest(),'docx_sha256':sha256(OUT.read_bytes()).hexdigest(),
            'changed_parts':sorted(changed),'preserved_parts':len(original)-len(changed-{ 'word/footerGuideLandscape.xml'}),
            'existing_numbering_preserved':True,'style_definitions':425,'logical_pages':len(content['pages']),
            'components_expected':sorted(public),'components_shown':sorted(marked),'component_instances':coverage,
            'native_word_review':'pending','pdf_export':'pending'}
    (TMP/'build-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(OUT)
if __name__=='__main__':main()
