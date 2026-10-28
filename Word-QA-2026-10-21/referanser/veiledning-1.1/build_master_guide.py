"""Build the versioned Norwegian master-template user guide as a PDF.

Run with the bundled Codex Python runtime. Content is edited in JSON, and
the manifest ties each guide revision to the exact master used for review.
"""
from pathlib import Path
import json
import hashlib
from html import escape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]
CONTENT = ROOT / 'docs/brukerveiledning-master.json'
OUT = ROOT / 'output/pdf/Virke_mastermal_brukerveiledning.pdf'
FONT_DIR = ROOT / 'Virke-revidert-2026-10-21/Fonter'


def main():
    content = json.loads(CONTENT.read_text())
    OUT.parent.mkdir(parents=True, exist_ok=True)
    pdfmetrics.registerFont(TTFont('Clan', str(FONT_DIR / 'ClanOffc-News.ttf')))
    pdfmetrics.registerFont(TTFont('ClanBold', str(FONT_DIR / 'ClanOffc-Bold.ttf')))
    pdfmetrics.registerFontFamily('Clan', normal='Clan', bold='ClanBold', italic='Clan', boldItalic='ClanBold')
    ink = HexColor('#111111'); muted = HexColor('#626262'); pink = HexColor('#C9007F')
    styles = {
        'body': ParagraphStyle('body', fontName='Clan', fontSize=10.1, leading=16.2, textColor=ink, spaceAfter=8),
        'intro': ParagraphStyle('intro', fontName='Clan', fontSize=11.1, leading=17.8, textColor=ink, spaceAfter=17),
        'title': ParagraphStyle('title', fontName='ClanBold', fontSize=23, leading=28, textColor=ink, spaceAfter=16),
        'heading': ParagraphStyle('heading', fontName='ClanBold', fontSize=12.2, leading=16, spaceBefore=0, spaceAfter=8, keepWithNext=True),
        'small': ParagraphStyle('small', fontName='Clan', fontSize=8.2, leading=13.1, textColor=muted, spaceAfter=8),
        'table': ParagraphStyle('table', fontName='Clan', fontSize=9.2, leading=14.7, textColor=ink),
        'th': ParagraphStyle('th', fontName='ClanBold', fontSize=9.2, leading=14.7, textColor=ink),
        'kicker': ParagraphStyle('kicker', fontName='ClanBold', fontSize=8.8, leading=12, textColor=pink, spaceAfter=12),
        'step': ParagraphStyle('step', fontName='Clan', fontSize=10.1, leading=16.2, textColor=ink, leftIndent=19, firstLineIndent=0, bulletIndent=0, spaceAfter=8),
    }
    width = A4[0]-104
    def para(t, kind='body'):
        return Paragraph(t.replace('\n','<br/>'), styles[kind])
    story=[]
    for page_i,page in enumerate(content['pages']):
        if page_i:story.append(PageBreak())
        story.extend([para(page['kicker'],'kicker'),para(page['title'],'title'),para(page['intro'],'intro')])
        for block in page['blocks']:
            kind=block['type']
            if kind in ('heading','paragraph','small'):
                if kind == 'heading':
                    # Explicit spacing survives ReportLab keep-with-next groups.
                    if story: story[-1].spaceAfter = 0
                    story.append(Spacer(1, 22.5))
                story.append(para(block['text'], 'body' if kind=='paragraph' else kind))
            elif kind in ('steps','bullets'):
                for i,item in enumerate(block['items'],1):
                    story.append(Paragraph(item, styles['step'],bulletText=f'{i}.' if kind=='steps' else '\u2022'))
            elif kind=='note':
                rows=[[para(block['title'],'th')],[para(block['text'],'table')]]
                t=Table(rows,colWidths=[width-2],hAlign='LEFT')
                t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),HexColor('#F9F2F6')),('BOX',(0,0),(-1,-1),0.5,HexColor('#E5D5DF')),('LEFTPADDING',(0,0),(-1,-1),12),('RIGHTPADDING',(0,0),(-1,-1),12),('TOPPADDING',(0,0),(-1,0),10),('BOTTOMPADDING',(0,0),(-1,0),4),('TOPPADDING',(0,1),(-1,-1),0),('BOTTOMPADDING',(0,1),(-1,-1),11)]))
                story.extend([t,Spacer(1,12)])
            elif kind=='table':
                rows=[[para(t,'th') for t in block['headers']]]+[[para(t,'table') for t in row] for row in block['rows']]
                t=Table(rows,colWidths=[width*f for f in block['widths']],repeatRows=1,hAlign='LEFT')
                t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),HexColor('#EEEEEE')),('VALIGN',(0,0),(-1,-1),'TOP'),('LINEBELOW',(0,0),(-1,-1),0.45,HexColor('#D1D1D1')),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),9),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
                story.extend([t,Spacer(1,12)])
            elif kind=='sources':
                story.extend([Spacer(1,5),para('Mer hjelp fra Microsoft','small')])
                links=[f'<link href="{escape(url,quote=True)}" color="#626262"><u>{escape(label)}</u></link>' for label,url in block['items']]
                story.append(para('<br/>'.join(links),'small'))
    total=len(content['pages'])
    def decorate(c,doc):
        if doc.page <= total:
            bookmark=f'guide-{doc.page}'
            c.bookmarkPage(bookmark)
            c.addOutlineEntry(content['pages'][doc.page-1]['title'],bookmark,level=0)
        c.saveState();c.setFillColor(muted);c.setFont('Clan',8)
        c.drawString(52,A4[1]-29,'Virke  /  Word')
        c.drawRightString(A4[0]-52,A4[1]-29,f"Brukerveiledning {content['version']}")
        c.drawString(52,27,f"{content['date']}  |  Virke mastermal")
        c.drawRightString(A4[0]-52,27,f'{doc.page} / {total}')
        c.restoreState()
    doc=SimpleDocTemplate(str(OUT),pagesize=A4,rightMargin=52,leftMargin=52,topMargin=59,bottomMargin=51,title=content['title'],author='Virke',subject='Bruk av mastermal i Microsoft Word',pageCompression=1)
    doc.build(story,onFirstPage=decorate,onLaterPages=decorate)
    actual_pages=len(PdfReader(str(OUT)).pages)
    if actual_pages != total:
        raise ValueError(f'Layout overflow: expected {total} pages, got {actual_pages}. Adjust layout and rebuild before delivery.')
    master=ROOT/content['master']
    manifest={
        'guide_version':content['version'],'date':content['date_iso'],
        'content_source':str(CONTENT.relative_to(ROOT)),
        'builder':'scripts/virke-build/build_master_guide.py',
        'pdf':str(OUT.relative_to(ROOT)),
        'master':content['master'],
        'master_sha256':hashlib.sha256(master.read_bytes()).hexdigest(),
        'pdf_sha256':hashlib.sha256(OUT.read_bytes()).hexdigest(),
        'change_note':content['change_note'],
        'revision_rule':'Oppdater veiledningen ved relevante endringer i stiler, felter, Autotekst, tabeller, nummerering, TOC, diagrammer, fonter, navn eller distribusjon.',
        'qa_status':'awaiting_render_review'
    }
    (ROOT/'docs/brukerveiledning-master-manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
    print(OUT)

if __name__=='__main__':main()
