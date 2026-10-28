from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from pypdf import PdfReader
import hashlib,json,shutil,subprocess,sys
root=Path('/Users/afg/Nettsteder/virke-word');qa=root/'Word-QA-2026-10-27';dest=root/'Virke-revidert-2026-10-27'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
master=dest/'Virke_master_test.dotx';docx=root/'output/docx/Virke_mastermal_brukerveiledning.docx';pdf=qa/'testkopier/Komponentguide-2.1-final.pdf'
record={'date':'2026-10-27','master_sha256':sha(master),'docx_sha256':sha(docx),'pdf_sha256':sha(pdf),'native_word_review':'passed','pdf_pages_reviewed':list(range(1,19)),'docx_render_pages_reviewed':list(range(1,19)),'native_word_pages_reviewed':list(range(1,19)),'native_review_evidence':'skjermbilder/guide-01.png through guide-18.png; final revised page: guide-final-18.png','final_revision_scope':'Pages 1–17 pixel-identical to previously reviewed Word export; page 18 revised and reviewed again in Word. All 18 final PDF and DOCX-render pages inspected.','pixel_comparison':'guide-final-pixel-comparison.json','pdf_export':'Microsoft Word for Mac, Save As PDF, Best for printing','components_shown':21,'observed_issues':[]}
review=qa/'guide-review-final.json';review.write_text(json.dumps(record,ensure_ascii=False,indent=2)+'\n')
subprocess.run([sys.executable,str(root/'scripts/virke-build/build_master_guide.py'),'--word-pdf',str(pdf),'--review-record',str(review)],check=True)
shutil.copy2(docx,dest/docx.name);shutil.copy2(root/'output/pdf/Virke_mastermal_brukerveiledning.pdf',dest/'Virke_mastermal_brukerveiledning.pdf')
# Retain exact functional-test scope; the baseline contains the all-block insertion test.
native=json.loads((qa/'native-forsok.json').read_text());native['status']='completed_with_documented_limitations'
native['tests'] += [
 {'id':'TABLE-GAP-3','file':'testkopier/27okt-fersk-master-v3.docx','result':'passed','observed':'Fresh master: insert two-column block then type immediately. Spacer remains exact 18 pt, followed by Normal continuation. Earlier v2 failure resolved.'},
 {'id':'MODULE-GAP-3','file':'testkopier/27okt-fersk-master-v3.docx','result':'passed_representative_insertion','observed':'Signature inserted, then immediate typing: exact 9 pt spacer retained. Same spacer pattern audited for summary and agenda blocks; their rendered examples reviewed.'},
 {'id':'HEADINGS','file':'testkopier/27okt-fersk-master-v3.docx','result':'passed','observed':'Long two-line H1, H1 after page break and repeated H3 reviewed. H1 22.5/13.5 pt and H3 18/9 pt in styles and final package audit.'},
 {'id':'FORM-FIELDS','file':'testkopier/27okt-fersk-master-v3.docx','result':'passed','observed':'Long name entered in first field; Tab to next field and second name entered. Fixed labels preserved, entered text black.'},
 {'id':'TOC-ACCESS-FINAL','file':'testkopier/27okt-fersk-master-v3.docx','result':'passed_alternative_access','observed':'Virke > Insert TOC opens native dialog, From template and three levels. Insert, modify heading, full field update and command-click navigation passed. Native gallery does not list custom entry.'},
 {'id':'EXAMPLES','result':'passed_visual_review','observed':'All 17 examples (18 pages) reviewed in Word using copies. Both landscape documents one page after removing an unnecessary end spacer. Final adjacent-table gaps rechecked in examples 09, 13 and 16. No opening prompts after detachment.'},
 {'id':'DETACH','result':'passed','observed':'All 17 final DOCX: attachedTemplate and its relationship removed; linkStyles absent; updateFields false; zero external relationships. Word copies opened without field or missing-template prompts.'},
 {'id':'GUIDE','result':'passed','observed':'All 21 actual components present; all 18 Word pages reviewed. Final PDF exported in Word, all pages inspected; page numbers 1–18 correct. Final DOCX render also inspected.'}
]
native['tests']=list({test['id']:test for test in native['tests']}.values())
native['baseline_regression_reference']='Virke-med-komponentguide-2026-10-21/TESTRESULTATER.txt'
native['coverage_note']='All 21 blocks were inserted in the earlier baseline test. This revision repeats changed behaviors with representative native insertions and reviews every component in the rebuilt guide. It does not claim 21 separate new insertion tests on 27 October.'
(qa/'native-forsok.json').write_text(json.dumps(native,ensure_ascii=False,indent=2)+'\n')
krav=json.loads((qa/'krav.json').read_text());krav['status']='completed_with_documented_limitations'
res={
 'H1':('passed','Stilaudit i master og alle 17 eksempler; ny H1, flerlint H1 og H1 etter sideskift i fersk masterkopi; eksempler 14/15 og guide.'),
 'H3':('passed','18 pt før og 9 pt etter i stilaudit; gjentatte H3 prøvd i fersk masterkopi og visuelt kontrollert.'),
 'TABLE-HEADER':('passed','Vanlig Sett inn > Tabell 3x3; utfylte overskrifter og ny kolonne forble fete. Ny datarad beholdt normal skrift. Lange tekster brytes.'),
 'CELL-BALANCE':('passed','Vanlige celler 9 pt topp / 3 pt bunn; særskilte sum- og signaturoppsett bevart. Alle eksempelsider og 21 guidekomponenter kontrollert.'),
 'TABLE-GAP':('implemented_with_scope_limit','18 pt etter ferdige Virke-tabellblokker bekreftet ved faktisk innsetting og umiddelbar skriving. 9 pt før i statiske eksempler; 18 pt mellom tilstøtende datatabeller. Vilkårlig vanlig Sett inn-tabell arver ikke automatisk en ytre tabellavstand fra en tabellstil.'),
 'MODULE-GAP':('passed_representative_native_test','9 pt etter Signatur bekreftet ved innsetting og skriving. Samme konstruksjon auditert for Saksliste og beløpsoppsummering; alle berørte komponenter gjennomgått i Word-guiden.'),
 'DISCOVERABILITY':('passed_mac_workflow','Virke-blokker synlig under Hjem og Virke. Knappen åpner dialogen; fanen Autotekst og Søk i master gir tilgang til blokkene. Ingen makroer.'),
 'TOC-ACCESS':('implemented_with_mac_gallery_limit','Synlig Virke-knapp åpner dialog med Fra mal/tre nivåer. Innsetting, ny overskrift, full oppdatering og navigasjon prøvd. Egen gallerioppføring vises ikke i Mac-galleriet etter forsøk med begge oppføringstyper.'),
 'BULLETS':('default_not_achieved_style_fallback_verified','Bindestrek+mellomrom gav svart strek. Vanlig punktlisteknapp gav svart prikk i nytt dokument. Punktliste-stilen gir rosa rund prikk og svart tekst og fortsetter med Enter. Ingen endring av brukerens globale Word-innstillinger.')}
for x in krav['requirements']:x['status'],x['evidence']=res[x['id']]
krav['detachment']={'status':'passed','evidence':'17 DOCX uten malrelasjon, ekstern relasjon eller automatisk stil-/feltoppdatering; åpnet uten varsel i Word.'}
krav['verification_scope']=native['coverage_note']
(qa/'krav.json').write_text(json.dumps(krav,ensure_ascii=False,indent=2)+'\n')
report='''Virke – testresultater 27. oktober 2026

OMFANG OG METODE
Master DOTX, 17 eksempeldokumenter / 18 sider, og komponentguide 2.1
med 21 faktiske redigerbare byggeblokker. Testendringer er gjort i kopier
med Codex sin innebygde datamaskinbruk direkte i Microsoft Word for Mac.
Supplerende DOCX-render er brukt til layoutkontroll. Leveransens PDF er
alltid eksportert fra Word med Best for utskrift.

NYE FUNKSJONSTESTER
• Ny tabell via vanlig Sett inn > Tabell: alle overskriftsceller fete,
  også etter utfylling og ny kolonne. Ny rad har normal skrift. Lang
  norsk tekst brytes i cellen uten klipping.
• Ferskt dokument fra revidert master: tabellblokk satt inn og tekst
  skrevet umiddelbart etterpå. 18 pt luft bevares; teksten bruker Normal.
• Signaturblokk satt inn med umiddelbar videre skriving. 9 pt luft
  bevares. Samme konstruksjon brukt for saksliste og beløpsoppsummering;
  disse er kontrollert som faktiske komponenter i guiden.
• Lang H1 på to linjer, H1 etter sideskift og gjentatte H3 kontrollert.
  Stilverdier: H1 22,5 pt før / 13,5 pt etter; H3 18 pt før / 9 pt etter.
• Utfyllingsmal: langt navn skrevet i første felt, Tab til neste felt.
  Faste etiketter blir stående, utfylt tekst er svart.
• Ny innholdsfortegnelse via Virke-knappen, Fra mal og tre nivåer.
  Overskrift endret, full feltoppdatering utført og lenkehopp kontrollert.
• Virke-blokker under Hjem og Virke åpner Autotekst-dialogen på Mac.
• Punktliste-stilen gir rosa runde punkter med svart tekst og fortsetter
  med Enter. Words standardknapp og bindestrek+mellomrom gir ikke
  automatisk den rosa varianten. Se Notat_til_Emil.txt.

VISUELL KONTROLL
Alle 17 eksempler gjennomgått i Word: marger, tabeller, utfyllingsfelt,
nummerering, diagrammer, innholdsfortegnelse og sidetall. Eksempel 01
har to sider; alle de andre har én. Unødige blanke sluttsider i de to
liggende dokumentene er fjernet. Avstand mellom tilstøtende tabeller
rettet og etterkontrollert i 09, 13 og 16. Alle 18 eksempelsider er også
inspisert i supplerende render. Ingen observerte klippe- eller overlappsfeil.

Alle 18 guidesider gjennomgått i Word. Alle 18 sider i den endelige
DOCX-renderen og Word-PDF-en visuelt inspisert. Sidetall 1–18 av 18,
moduloversikt og tre levende TOC-henvisninger er korrekte. Liggende
sider 10/11 og diagrammer på 15/16 er bevart. Ved den siste revisjonen
var side 1–17 pikselidentiske med tidligere kontrollert Word-eksport;
side 18 ble revidert og kontrollert på nytt direkte i Word.

TESTDEKNING OVER TID
Alle 21 blokker ble satt inn i Word i forrige mastertest. Denne runden
kontrollerer alle 21 i komponentguiden og gjentar innsetting/redigering
for de endrede funksjonene med representative blokker. Diagramdata,
bilder, tema og innebygde regneark er byteidentiske med forrige
leveranse. Den tidligere beståtte Excel-dataredigeringen er ikke
utført på nytt i denne runden.

FRIKOBLING OG PAKKEKONTROLL
Alle 17 eksempeldokumenter har innebygd formatering og ingen nødvendig
maltilknytning. Attached-template-relasjonen er fjernet, automatisk
stiloppdatering er av og automatisk feltoppdatering ved åpning er av.
Ingen eksterne relasjoner finnes i de leverte eksemplene. Kopiene åpnet
uten feltvarsel etter frikoblingen. Veiledningen er også selvstendig.

Alle XML-deler kan parses. Alle 21 offentlige Autotekst-komponenter er
bevart; TOC-gallerialiaset teller ikke som en ekstra komponent.
Originaler og før-versjoner er bevart. Ingen QA-tekst eller testkopier
inngår i leveransen. Ingen commit eller push er utført.

AVGRENSNINGER
Windows og Word på nett er ikke testet. Tilgang via Mac-galleri,
standard punktlisteknapp og ytre luft ved vilkårlig ny tabell har
begrensningene beskrevet i Notat_til_Emil.txt. De tidligere kosmetiske
TOC-avvikene består. Eksempel 08 viderefører det manuelt skrevne 5.1
fra designgrunnlaget; automatisk flernivånummerering finnes i Lovtekst
og eksempel 17. De 40 Styres-konverteringene er ikke del av denne pakken.
'''
(dest/'TESTRESULTATER.txt').write_text(report)
# Replace stale working-contract history with current build contract. Before-version is preserved.
(root/'tmp/master-guide-docx/artifact.md').write_text(f'''# Virke komponentguide 2.1 – gjeldende dokumentkontrakt

Referanse: {master}
SHA-256: {sha(master)}
Kilde: docs/brukerveiledning-master.json. Bygger: scripts/virke-build/build_master_guide_docx.py.

18 sider. Stående A4 med masterens geometri, egne liggende sider 10 og 11.
Alle 21 offentlige Autotekst-komponenter vises som faktiske redigerbare elementer.
H1: 22,5 pt før / 13,5 pt etter. H3: 18 pt før / 9 pt etter.
Vanlige celler: 9 pt topp / 3 pt bunn. Egne spacer-stiler bevarer 18 pt etter datatabell
og 9 pt etter modul; guideforklaringen bruker avsnittsavstand i stedet for en ekstra tom linje.
Rosa punkttegn via Punktliste. Ingen designavvik eller kosmetiske sluttinstrukser i guiden.

Stiler, Autotekst, diagrammer, medier, tema og innebygde regneark fra masteren bevares.
Nye listeinstanser, modulbokmerker og levende sidehenvisninger tilføyes av byggeren.
Innholdsfortegnelsens feltresultat importeres fra native Word uten å omskrive andre pakkedeler.
DOCX og PDF er selvstendige; ingen nødvendig ekstern malrelasjon.

Verifisert 27. oktober: alle Word-sider og alle endelige DOCX-/PDF-rendersider kontrollert.
PDF eksportert fra Microsoft Word. Kontrollpost: Word-QA-2026-10-27/guide-review-final.json.
Før-kontrakten er bevart i Word-QA-2026-10-27/for-versjon/artifact.md.
''')
bpath=root/'tmp/master-guide-docx/build-report.json';br=json.loads(bpath.read_text());br.update(native_word_review='passed',pdf_export='Microsoft Word for Mac',review_record=str(review.relative_to(root)));bpath.write_text(json.dumps(br,ensure_ascii=False,indent=2)+'\n')
# Guide is detached and has all real module bookmarks.
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
with ZipFile(docx) as z:
 settings=E.fromstring(z.read('word/settings.xml'));assert not settings.findall('w:attachedTemplate',ns)
 document=E.fromstring(z.read('word/document.xml'));marks=document.xpath('//w:bookmarkStart/@w:name',namespaces=ns)
 assert len(set(x for x in marks if x.startswith('GuideModule')))==21
 for n in z.namelist():
  if n.endswith('.rels'):assert not any(x.get('TargetMode')=='External' for x in E.fromstring(z.read(n)))
reader=PdfReader(pdf);assert len(reader.pages)==18
links=sum(len(p['/Annots'].get_object()) if '/Annots' in p else 0 for p in reader.pages);assert links>=24,links
content_files=sorted(p for p in dest.rglob('*') if p.is_file() and not p.name.startswith('.') and p.name!='LEVERANSE-MANIFEST.json')
manifest={'date':'2026-10-27','guide_version':'2.1','master_sha256':sha(master),'guide_docx_sha256':sha(docx),'guide_pdf_sha256':sha(pdf),'platform':'Word for Mac','example_documents':17,'example_pages':18,'guide_pages':18,'components':21,'guide_pdf_annotations':links,'files':[{'file':str(p.relative_to(dest)),'sha256':sha(p),'bytes':p.stat().st_size} for p in content_files]}
(dest/'LEVERANSE-MANIFEST.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
zip_path=dest.with_suffix('.zip');assert not zip_path.exists(),'Preserve existing ZIP; choose new name before running again'
with ZipFile(zip_path,'w',ZIP_DEFLATED,compresslevel=9) as z:
 for p in sorted(dest.rglob('*')):
  if p.is_file() and not p.name.startswith('.') and not p.name.startswith('~$'):z.write(p,p.relative_to(root))
with ZipFile(zip_path) as z:
 assert z.testzip() is None
 names=z.namelist();assert sum(n.endswith('.docx') for n in names)==18
 assert sum(n.endswith('.dotx') for n in names)==1
 assert sum(n.endswith('.pdf') for n in names)==1
 assert sum(n.endswith('.ttf') for n in names)==2
 for rec in manifest['files']:assert hashlib.sha256(z.read(dest.name+'/'+rec['file'])).hexdigest()==rec['sha256']
(qa/'delivery-result.json').write_text(json.dumps({'zip':str(zip_path),'sha256':sha(zip_path),'bytes':zip_path.stat().st_size,'integrity':'passed','manifest':str(dest/'LEVERANSE-MANIFEST.json')},indent=2)+'\n')
print(json.dumps({'zip':str(zip_path),'bytes':zip_path.stat().st_size,'files':len(names),'sha256':sha(zip_path),'pdf_links':links},indent=2))
