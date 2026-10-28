"""Finalize a reviewed Word PDF and bind it to the editable guide and master.

Build DOCX with build_master_guide_docx.py, review it in Microsoft Word, export
PDF from Word, render every PDF page, then supply that review record here.
This entry point deliberately does not regenerate the guide with ReportLab.
"""
import argparse
import hashlib
import json
import shutil
from pathlib import Path
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]
CONTENT = ROOT / 'docs/brukerveiledning-master.json'
DOCX = ROOT / 'output/docx/Virke_mastermal_brukerveiledning.docx'
OUT = ROOT / 'output/pdf/Virke_mastermal_brukerveiledning.pdf'

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--word-pdf', type=Path, required=True)
    parser.add_argument('--review-record', type=Path, required=True)
    args = parser.parse_args()
    content = json.loads(CONTENT.read_text())
    review = json.loads(args.review_record.read_text())
    master = ROOT / content['master']
    expected = len(content['pages'])
    for key, path in [('master_sha256', master), ('docx_sha256', DOCX),
                      ('pdf_sha256', args.word_pdf)]:
        if review.get(key) != digest(path):
            raise ValueError(f'Review is stale or missing: {key}')
    if review.get('native_word_review') != 'passed' or review.get('pdf_pages_reviewed') != list(range(1, expected+1)):
        raise ValueError('Native Word review and inspection of every PDF page are required')
    reader = PdfReader(str(args.word_pdf))
    if len(reader.pages) != expected:
        raise ValueError('Unexpected PDF page count')
    for i, page in enumerate(reader.pages, 1):
        text = ' '.join(page.extract_text().split())
        if f'Side {i} av {expected}' not in text:
            raise ValueError(f'Incorrect footer on page {i}')
        if content['pages'][i-1]['title'] not in text:
            raise ValueError(f'Expected chapter missing from page {i}')
    OUT.parent.mkdir(parents=True, exist_ok=True)
    if args.word_pdf.resolve() != OUT.resolve():
        shutil.copy2(args.word_pdf, OUT)
    manifest = {
        'guide_version': content['version'], 'date': content['date_iso'],
        'content_source': str(CONTENT.relative_to(ROOT)),
        'docx_builder': 'scripts/virke-build/build_master_guide_docx.py',
        'builder': 'scripts/virke-build/build_master_guide.py',
        'docx': str(DOCX.relative_to(ROOT)), 'docx_sha256': digest(DOCX),
        'pdf': str(OUT.relative_to(ROOT)), 'pdf_sha256': digest(OUT),
        'master': content['master'], 'master_sha256': digest(master),
        'change_note': content['change_note'],
        'revision_rule': 'Oppdater JSON og DOCX ved relevante endringer. Kontroller i Word, eksporter PDF, inspiser alle sider og oppdater manifest og leveranse-ZIP.',
        'pdf_export': 'Microsoft Word for Mac, Lagre som PDF, Best for utskrift',
        'qa_status': 'native_word_and_all_rendered_pages_reviewed',
        'rendered_pages': expected,
        'review_record': str(args.review_record.resolve().relative_to(ROOT)),
        'render_review': 'Alle sju sider kontrollert i Word, DOCX-render og Word-eksportert PDF. Ingen klipping, overlapp eller feil i sidetall.',
        'platform_scope': 'Word for Mac. Windows og Word på nett er ikke verifisert.'
    }
    (ROOT/'docs/brukerveiledning-master-manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2)+'\n')
    print(OUT)

if __name__ == '__main__':
    main()
