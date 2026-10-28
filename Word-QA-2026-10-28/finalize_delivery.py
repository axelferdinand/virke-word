"""Package the table revision only after native Word and page reviews."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib, json, shutil, subprocess, sys
from pypdf import PdfReader

root = Path(__file__).resolve().parents[1]
qa = root / 'Word-QA-2026-10-28'
dest = root / 'Virke-revidert-2026-10-28'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
master = dest / 'Virke_master_test.dotx'
docx = root / 'output/docx/Virke_mastermal_brukerveiledning.docx'
pdf = qa / 'testkopier/Komponentguide-2.2-slutt.pdf'
assert sha(docx) == sha(qa/'testkopier/Komponentguide-2.2-slutt.docx')
assert sha(master) == '3e4f4d5d2fb37e62f7a91df24d6b9d85acd1071dff78a87bc75eee6eecedae5c'
for n in range(1, 19):
    assert (qa/f'skjermbilder/guide-side-{n:02}.png').exists()
    assert (qa/f'guide-word-pdf/page-{n:02}.png').exists()

review = qa / 'guide-review-final.json'
review.write_text(json.dumps({
    'date': '2026-10-28', 'master_sha256': sha(master),
    'docx_sha256': sha(docx), 'pdf_sha256': sha(pdf),
    'native_word_review': 'passed',
    'native_word_pages_reviewed': list(range(1,19)),
    'pdf_pages_reviewed': list(range(1,19)),
    'docx_render_pages_reviewed': list(range(1,19)),
    'native_review_evidence': 'skjermbilder/guide-side-01.png through guide-side-18.png',
    'pdf_review_evidence': 'guide-word-pdf-review/review-01.png through review-17.png (pairs)',
    'final_docx_render_review': 'Page 4 inspected after native TOC cache import; the other 17 pages are pixel-identical to the fully reviewed guide-render-v2.',
    'pixel_comparison': 'guide-final-pixel-comparison.json',
    'pdf_export': 'Microsoft Word for Mac, Save As PDF, Best for printing',
    'components_shown': 21, 'observed_issues': []
}, ensure_ascii=False, indent=2)+'\n')
subprocess.run([sys.executable, str(root/'scripts/virke-build/build_master_guide.py'),
                '--word-pdf', str(pdf), '--review-record', str(review)], check=True)
shutil.copy2(docx, dest/docx.name)
shutil.copy2(root/'output/pdf/Virke_mastermal_brukerveiledning.pdf', dest/'Virke_mastermal_brukerveiledning.pdf')

(dest/'START_HER.txt').write_text('''VIRKE – MASTER OG KOMPONENTGUIDE
Revidert 28. oktober 2026. Veiledning 2.2.

1. Installer de to Clan-fontene fra Fonter og start Word på nytt.
2. Les Virke_mastermal_brukerveiledning.pdf. DOCX-utgaven er redigerbar
   og viser alle 21 komponentene med praktiske forklaringer.
3. Åpne Virke_master_test.dotx for å opprette et nytt dokument med
   masterens stiler og Virke-blokker. Lagre dokumentet som DOCX.
4. Bruk kopier av de nummererte eksempeldokumentene når oppsettet passer.
   De 17 DOCX-filene viser til sammen 18 eksempelsider.

Eksemplene er selvstendige og kan flyttes til Virkes Microsoft 365-arkiv.
Formateringen ligger i hver fil. Det kreves ingen tilknytning til
masterfilens plassering, og automatisk stil- og feltoppdatering ved
åpning er slått av. Oppdater innholdsfortegnelsen etter innholdsendringer.

Bruk Word på skrivebordet. Arbeidsmåtene er kontrollert i Word for Mac.
Fontene må installeres; de lastes ikke automatisk fra leveransemappen.
Master, fonter og veiledning oppbevares samlet for videre vedlikehold.

Notat_til_Emil.txt og TESTRESULTATER.txt dokumenterer denne revisjonen.
De 40 Styres-malene er neste leveransefase og inngår ikke i denne ZIP-en.
''')

(dest/'Notat_til_Emil.txt').write_text('''Til Emil – tabellrevisjon 28. oktober 2026

Du hadde rett i at forrige tabellretting var utilstrekkelig. Normal-stilen
overstyrte tabellstilens skriftstørrelse og avsnittsavstand. Dette er nå
rettet i mastermalen, de lagrede Autotekst-blokkene og alle 17 eksemplene.
Vanlig brødtekst beholder 9 pkt, mens tabellstilen får styre 7,5 pkt,
linjeavstand, cellemarger og fet overskriftsrad.

Den reviderte versjonen er prøvd direkte i Word for Mac med vanlig
Sett inn > Tabell fra Normal, utfylling av første overskriftscelle,
ny kolonne, ny rad med Tab, lang norsk tekst og lagring/gjenåpning.
Disse prøvene ga 7,5 pkt, fete overskrifter i alle kolonner og vanlig
skrift i dataradene. En tabell satt inn fra masterens Autotekst ble
også prøvd, inkludert utfylling av første overskriftsfelt.

Cellebalansen er justert mot Figma med kompensasjon for hvordan Word
plasserer Clan-fontens linjeboks. Standardcellene bruker 6 pkt på sidene,
5,25 pkt over og 8,25 pkt under. Ulike tekniske topp-/bunnverdier gir
bedre optisk balanse i Word. Lange innholdsrader er toppjustert der
Figma viser dette. «Store celler» beholder bevisst 9 pkt i den første
kontrolltekstkolonnen, slik denne varianten er spesifisert i Figma.

Alle 17 eksempler og alle 18 sider i komponentguide 2.2 er gjennomgått
i Word. Guiden viser fortsatt alle 21 redigerbare byggeblokker. PDF-en
er eksportert fra Word og alle sider er inspisert.

Avgrensninger som videreføres fra forrige runde:
• Rosa punkttegn velges med stilen Punktliste. Words standardknapp
  og automatisk liste fra bindestrek kan bruke egne standarder.
• På Mac åpnes Virke-fortegnelsen via Virke-blokker eller Virke-knappen
  for innholdsfortegnelse. Egen oppføring i standardgalleriet ble ikke
  oppnådd i forrige test.
• Fast luft etter en tabell følger de ferdige Virke-blokkene. En helt
  ny standardtabell får tabellformateringen, men ingen garanti om
  automatisk 18 pkt avstand til det etterfølgende avsnittet.
• Tidligere kosmetiske forskjeller i innholdsfortegnelsens øverste
  linje og sidetallsfarge på første nivå inngår ikke i denne tabellrettingen.

Kontrollen gjelder Word for Mac. Windows og Word på nett er ikke testet.
De beskrevne tabellfeilene er ikke avskrevet som uunngåelige Word-feil;
de er rettet og bestått i de konkrete prøvene over.
''')

(dest/'TESTRESULTATER.txt').write_text('''VIRKE – TESTRESULTATER 28. OKTOBER 2026

OMFANG
1 DOTX, 17 eksempler / 18 sider og komponentguide 2.2 med 21 faktiske
redigerbare byggeblokker. Testendringer er gjort i kopier, direkte i
Microsoft Word for Mac med innebygd datamaskinbruk. Windows og Word på
nett er ikke testet. De 40 Styres-konverteringene er ikke startet.

HVA DENNE RAPPORTEN ERSTATTER
Konklusjonen om cellebalanse i rapporten fra 27. oktober var feil.
Den tidligere kontrollen fanget ikke at Normal overstyrte tabellstilen.
Denne revisjonen erstatter de tidligere påstandene om tabellstørrelse,
cellebalanse og formatering ved vanlig tabellinnsetting.

ÅRSAK OG RETTING
Normal hadde eksplisitt 9 pkt og avsnittsformatering med høyere prioritet
enn tabellstilen. Disse grunnverdiene er flyttet til dokumentstandardene.
Normal ser fortsatt lik ut utenfor tabeller, og tabellstilen kan nå styre
7,5 pkt, 12 pkt linjeavstand, null avsnittsavstand og cellemarger.
Fet skrift i overskriftsraden har én eier: tabellstilens første-rad-regel.
Konkurrerende fet/ikke-fet-overstyringer er fjernet fra tabelloverskrifter.
Top/bunn er optisk balansert med 5,25/8,25 pkt; sidene bruker 6 pkt.
Overflødige minste-radhøyder er fjernet. «Store celler» har en bevisst
9 pkt kontrolltekstkolonne og toppjusterte lange datarader.

NYE FUNKSJONSPRØVER I WORD
Bestått:
• Før-versjon fra 27. oktober: 9 pkt ved vanlig tabellinnsetting gjenskapt.
• Nytt dokument fra revidert DOTX: vanlig Sett inn > Tabell 3x3 fra Normal
  gir 7,5 pkt; alle tre overskriftsceller fete, datarad vanlig skrift.
• Endelig eksempel 02: første overskriftsfelt fylt med «Oktober».
  Feltet er svart og fett i 7,5 pkt. Valg av Normal beholder tabellstørrelsen.
• Ny standardtabell i testkopi av eksempel 02, deretter lagring og
  gjenåpning: 7,5 pkt og fet første overskrift bevart.
• Ny kolonne i samme tabell: ny overskrift arver 7,5 pkt og fet skrift.
• Lang norsk tekst med æ, ø og å: automatisk linjebryting uten klipping.
• Tab i siste celle: ny fjerde datarad med 7,5 pkt og vanlig skrift.
• Nytt dokument fra endelig master: Autotekst «Tabell med 3 kolonner»
  satt inn. Første overskriftsfelt fylt ut; 7,5 pkt og fet Clan bekreftet.
• Veiledningens felt oppdatert i Word; tre TOC-lenker får riktige titler
  og side 3. Native feltresultat importert uten å omskrive øvrige deler.

VISUELL KONTROLL
Alle 17 eksempler kontrollert i Word, inkludert begge sider i eksempel
01. Øvrige eksempler er én side hver. Sideoppsett, tabeller, felter,
nummerering, innholdsfortegnelse og diagrammer er visuelt gjennomgått.
Alle 18 eksempelsider er også kontrollert i supplerende DOCX-render.
Eksempel 04 har én færre tom utfyllingsrad for å beholde én liggende
side etter korrigert tabellhøyde. Ingen utfylte opplysninger er fjernet.

Alle 18 guidesider kontrollert direkte i Word. Alle 18 sider i Word-PDF
inspisert. Sidetall 1–18 av 18, to liggende sider og 24 interne lenker
kontrollert. Alle 21 komponenter er med. Endelig DOCX-render: side 4
kontrollert etter feltoppdatering; øvrige 17 sider er pikselidentiske
med den allerede fullt kontrollerte versjonen fra denne revisjonen.
Ingen observerte klippe- eller overlappsfeil i sluttversjonene.

TESTDEKNING
Denne runden gjentar de endrede tabellfunksjonene og gjennomgår alle
komponenter visuelt. Den påstår ikke 21 separate nye Autotekst-innsettinger.
Alle blokkene ble satt inn i den tidligere baselinetesten; diagramdata
ble også redigert da. Excel-redigering er ikke gjentatt i denne runden.
Nummerering og øvrige moduler er bevart og visuelt regresjonskontrollert.

PAKKEKONTROLL
Master og alle 17 eksempler har samme korrigerte stilregler. Alle XML-
deler kan parses. Diagrammer, medier, tema og innebygde regneark er
byteidentiske med forrige leveranse. Alle 21 offentlige Autotekst-blokker
er bevart. Utfyllingsfelt er bevart bortsett fra den tomme raden i 04.
DOCX-filene er selvstendige, uten ekstern malrelasjon eller automatisk
stil-/feltoppdatering ved åpning. Word-kopiene åpnet uten mal-/feltvarsel.
Før-versjoner og QA-kopier er bevart utenfor leveransen.
ZIP-en inneholder ingen QA-tekst, testkopier eller midlertidige låsefiler.
Ingen commit eller push er utført. Se Notat_til_Emil.txt for avgrensninger.
''')

native = {
    'date': '2026-10-28', 'platform': 'Microsoft Word for Mac',
    'method': 'Built-in computer use, mutations restricted to QA copies',
    'status': 'passed_with_documented_platform_scope',
    'all_examples_reviewed': list(range(1,18)), 'example_pages': 18,
    'tests': [
        {'id': 'BASELINE-REPRO', 'result':'9 pt reproduced in the 27 October package'},
        {'id': 'NATIVE-INSERT', 'result':'7.5 pt all header/body cells; header bold and body regular in a fresh master document and example 02 copy'},
        {'id': 'REOPEN', 'result':'First header stays bold 7.5 pt after save/reopen', 'evidence':'sluttkontroll/02_Balance_scorecard.docx'},
        {'id': 'COLUMN', 'result':'Inserted header column bold 7.5 pt', 'evidence':'skjermbilder/slutt-kolonne-og-langtekst.png'},
        {'id': 'ROW', 'result':'Tab creates fourth row; 7.5 pt Clan body, bold off', 'evidence':'skjermbilder/slutt-ny-rad.txt'},
        {'id': 'AUTOTEXT', 'result':'Inserted final master 3-column block; first header filled; bold 7.5 pt', 'evidence':'skjermbilder/master-autotekst-header-verdier.txt'},
        {'id': 'GUIDE', 'result':'18 Word pages and 18 Word PDF pages reviewed; 21 actual components; native field update passed', 'evidence':'guide-review-final.json'}
    ]
}
(qa/'native-forsok.json').write_text(json.dumps(native, ensure_ascii=False, indent=2)+'\n')

contract = root/'tmp/master-guide-docx/artifact.md'
contract.write_text(contract.read_text().replace('18 planlagte sider', '18 kontrollerte sider').replace(
    'Status: native sluttkontroll og PDF-registrering pågår.',
    'Status: alle Word-sider og Word-PDF-sider visuelt kontrollert. Endelig DOCX-render kontrollert; 17 sider pikselidentiske med tidligere kontroll, oppdatert TOC-side 4 gjennomgått. Kontrollpost: Word-QA-2026-10-28/guide-review-final.json.'))
bp=root/'tmp/master-guide-docx/build-report.json'; b=json.loads(bp.read_text())
b.update(native_word_review='passed', pdf_export='Microsoft Word for Mac', review_record=str(review.relative_to(root)))
bp.write_text(json.dumps(b, ensure_ascii=False, indent=2)+'\n')

subprocess.run([sys.executable,str(qa/'audit_package.py')],check=True)
reader=PdfReader(pdf); links=sum(len(p['/Annots'].get_object()) if '/Annots' in p else 0 for p in reader.pages)
assert len(reader.pages)==18 and links==24
content_files=sorted(p for p in dest.rglob('*') if p.is_file() and not p.name.startswith(('.', '~$')) and p.name!='LEVERANSE-MANIFEST.json')
manifest={
    'date':'2026-10-28', 'guide_version':'2.2', 'master_sha256':sha(master),
    'guide_docx_sha256':sha(docx), 'guide_pdf_sha256':sha(pdf),
    'platform':'Word for Mac; Windows and web not verified',
    'example_documents':17,'example_pages':18,'guide_pages':18,'components':21,
    'guide_pdf_annotations':links,
    'files':[{'file':str(p.relative_to(dest)),'sha256':sha(p),'bytes':p.stat().st_size} for p in content_files]
}
(dest/'LEVERANSE-MANIFEST.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
zip_path=dest.with_suffix('.zip')
assert not zip_path.exists(), 'Preserve the existing ZIP before rebuilding'
with ZipFile(zip_path,'w',ZIP_DEFLATED,compresslevel=9) as z:
    for p in sorted(dest.rglob('*')):
        if p.is_file() and not p.name.startswith(('.', '~$')): z.write(p,p.relative_to(root))
with ZipFile(zip_path) as z:
    assert z.testzip() is None
    names=z.namelist()
    for ext,count in [('.docx',18),('.dotx',1),('.pdf',1),('.ttf',2)]:
        assert sum(n.endswith(ext) for n in names)==count,(ext,names)
    for rec in manifest['files']:
        assert hashlib.sha256(z.read(dest.name+'/'+rec['file'])).hexdigest()==rec['sha256']
result={'zip':str(zip_path),'sha256':sha(zip_path),'bytes':zip_path.stat().st_size,'integrity':'passed','files':len(names),'pdf_links':links}
(qa/'delivery-result.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps(result,indent=2))
