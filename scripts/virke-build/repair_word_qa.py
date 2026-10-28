"""Apply the Word QA corrections to a clean package, including AutoText parts.

Run with the bundled Python runtime. Input is the immutable pre-repair ZIP;
output is a staging directory, so interactive QA never modifies the deliverables.
"""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from io import BytesIO
import argparse
import json
import posixpath
from lxml import etree as E
from openpyxl import load_workbook, Workbook

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'w': W, 'r': R}
Q = lambda name: '{' + W + '}' + name


def xml(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def read_package(path):
    with ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def text(node):
    return ''.join(node.xpath('.//w:t/text()', namespaces=NS))


def repair_content(root):
    controls = rows = 0
    for table in root.iter(Q('tbl')):
        cells = table.findall('.//' + Q('tc'))
        # This identifies the shared labelled form, not ordinary data tables.
        is_form = bool(cells) and all(text(c).startswith('Navn: ') and c.find('.//' + Q('sdt')) is not None for c in cells)
        if is_form:
            for props in table.iter(Q('sdtPr')):
                for child in list(props):
                    if child.tag in (Q('temporary'), Q('text'), Q('richText')):
                        props.remove(child)
                E.SubElement(props, Q('text'))
                controls += 1
        # Preserve the content-determined height; new rows start small.
        for height in table.findall('.//' + Q('trHeight')):
            if int(height.get(Q('val'), '0')) >= 3000:
                height.set(Q('val'), '360')
                height.set(Q('hRule'), 'atLeast')
                rows += 1
    return controls, rows


def repair_styles(root):
    for style in root.xpath('w:style[@w:styleId="Listeavsnitt"]', namespaces=NS):
        style.find(Q('next')).set(Q('val'), 'Listeavsnitt')


def corrected_toc(document, tested):
    old = document.xpath('//w:sdt[w:sdtPr/w:docPartObj]', namespaces=NS)[0]
    new = deepcopy(tested.xpath('//w:sdt[w:sdtPr/w:docPartObj]', namespaces=NS)[0])
    for paragraph in list(new.find(Q('sdtContent'))):
        if 'QA-test' in text(paragraph):
            paragraph.getparent().remove(paragraph)
    old.getparent().replace(old, new)
    heading = next(p for p in document.findall('.//' + Q('p')) if text(p) == '1. Bakgrunn og Virksomhetens formål')
    anchor = new.find('.//' + Q('hyperlink')).get(Q('anchor'))
    ids = [int(x.get(Q('id'))) for x in document.iter(Q('bookmarkStart'))]
    bid = str(max(ids, default=0) + 1)
    heading.insert(1, E.Element(Q('bookmarkStart'), {Q('id'): bid, Q('name'): anchor}))
    heading.append(E.Element(Q('bookmarkEnd'), {Q('id'): bid}))


def clean_chart_workbook(raw):
    """Rebuild the small chart data grid, without broken Excel table metadata."""
    old = load_workbook(BytesIO(raw), data_only=False)
    new = Workbook()
    new.remove(new.active)
    for sheet in old:
        target = new.create_sheet(sheet.title)
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    target.cell(cell.row, cell.column, cell.value)
    out = BytesIO()
    new.save(out)
    return out.getvalue()


def repair_charts(data, recovered=None):
    # Word's own round trip has validated this chart markup and its workbooks.
    # Keep the document and glossary references (chart1/chart2) unchanged.
    # Only example 11 needs recovered chart markup. Other examples and the
    # master retain their original chart designs, series and values.
    source = recovered if recovered is not None else dict(data)
    for name in list(data):
        if name.startswith(('word/charts/', 'word/embeddings/')):
            del data[name]
    for name, content in source.items():
        if name.startswith(('word/charts/', 'word/embeddings/')):
            data[name] = clean_chart_workbook(content) if name.endswith('.xlsx') else content
    ct = E.fromstring(data['[Content_Types].xml'])
    for entry in list(ct):
        if entry.get('PartName', '').startswith(('/word/charts/', '/word/embeddings/')):
            ct.remove(entry)
    recovered_ct = E.fromstring(source['[Content_Types].xml'])
    for entry in recovered_ct:
        if entry.get('PartName', '').startswith(('/word/charts/', '/word/embeddings/')):
            ct.append(deepcopy(entry))
    # A chart embedded in the document and an AutoText chart need independent
    # package parts. Sharing a chart part across owners triggered Word's
    # unreadable-content recovery in example 11.
    for name, content in list(data.items()):
        if name.startswith(('word/charts/', 'word/embeddings/')):
            data[name.replace('word/', 'word/glossary/', 1)] = content
    for entry in recovered_ct:
        if entry.get('PartName', '').startswith(('/word/charts/', '/word/embeddings/')):
            duplicate = deepcopy(entry)
            duplicate.set('PartName', duplicate.get('PartName').replace('/word/', '/word/glossary/', 1))
            ct.append(duplicate)
    glossary_rels = E.fromstring(data['word/glossary/_rels/document.xml.rels'])
    for rel in glossary_rels:
        if rel.get('Type', '').endswith('/chart'):
            rel.set('Target', 'charts/' + posixpath.basename(rel.get('Target')))
    data['word/glossary/_rels/document.xml.rels'] = xml(glossary_rels)
    data['[Content_Types].xml'] = xml(ct)


def validate(data):
    for name, raw in data.items():
        if name.endswith(('.xml', '.rels')):
            root = E.fromstring(raw)
            if name.endswith('.rels'):
                scope = posixpath.dirname(posixpath.dirname(name))
                for rel in root:
                    if rel.get('TargetMode') == 'External':
                        continue
                    target = rel.get('Target', '')
                    resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(scope, target))
                    assert resolved in data, (name, resolved)
    glossary = E.fromstring(data['word/glossary/document.xml'])
    assert len(glossary.xpath('w:docParts/w:docPart[w:docPartPr/w:category/w:gallery[@w:val="autoTxt"]]', namespaces=NS)) == 20
    assert b'QA-test' not in data['word/document.xml']


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--before', type=Path, required=True)
    parser.add_argument('--recovered-chart-document', type=Path, required=True)
    parser.add_argument('--tested-toc-document', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)
    recovered = read_package(args.recovered_chart_document)
    tested = E.fromstring(read_package(args.tested_toc_document)['word/document.xml'])
    manifest = []
    with ZipFile(args.before) as archive:
        for name in archive.namelist():
            if not name.endswith(('.docx', '.dotx')):
                continue
            data = read_package(BytesIO(archive.read(name)))
            changes = {}
            for part in ('word/document.xml', 'word/glossary/document.xml'):
                root = E.fromstring(data[part])
                controls, rows = repair_content(root)
                if name.startswith('07_') and part == 'word/document.xml':
                    corrected_toc(root, tested)
                data[part] = xml(root)
                changes[part] = {'form_controls': controls, 'row_heights': rows}
            for part in ('word/styles.xml', 'word/glossary/styles.xml'):
                root = E.fromstring(data[part])
                repair_styles(root)
                data[part] = xml(root)
            repair_charts(data, recovered if name.startswith('11_') else None)
            validate(data)
            with ZipFile(args.out / name, 'w', ZIP_DEFLATED) as out:
                for part, raw in data.items():
                    out.writestr(part, raw)
            manifest.append({'file': name, 'changes': changes})
    (args.out.parent / 'rettinger.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    print(f'Patched and structurally verified {len(manifest)} files in {args.out}')


if __name__ == '__main__':
    main()
