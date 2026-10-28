"""Record the completed interactive checks and assemble review documentation."""
from pathlib import Path
from hashlib import sha256
import json, plistlib, shutil
ROOT=Path(__file__).resolve().parents[1];QA=ROOT/'Word-QA-2026-10-21';D=ROOT/'Virke-revidert-2026-10-21'
def write(p,s):p.write_text(s,encoding='utf-8')
def js(p,d):write(p,json.dumps(d,ensure_ascii=False,indent=2)+'\n')
with Path('/Applications/Microsoft Word.app/Contents/Info.plist').open('rb') as f:word_version=plistlib.load(f).get('CFBundleShortVersionString')
examples=[
('01_Protokoll_styremote.docx',2,'Begge sider gjennomgått. Tabell, signaturer, marger og Side 2 av 2 korrekte.'),
('02_Balance_scorecard.docx',1,'Begge tabeller kontrollert. Lang norsk tekst utvidet testkopien til fem sider; overskriftsrad gjentatt, ingen kolonneklipping.'),
('03_Utfyllingsskjema.docx',1,'Ni utfyllingsfelt og etiketter kontrollert visuelt. Tab/Skift+Tab og svart utfylt tekst ettertestet i masterens Utfyllingsmal.'),
('04_Aksjeeierbok_liggende.docx',1,'Liggende A4, ni kolonner og fete overskrifter innenfor margene.'),
('05_Protokoll_liggende_ni_kolonner.docx',1,'Liggende A4, hele tabellen og bunntekst kontrollert.'),
('06_Protokoll_Prototypen_AS.docx',1,'Brødtekst, tabell og rosa signaturfelt gjennomgått.'),
('07_Innholdsfortegnelse_og_avtale.docx',1,'Hele fortegnelsen oppdatert i Word. H1 og H2 inngår. Kun oppdatert feltresultat og tilhørende bokmerker overført til ren leveranse.'),
('08_Avtalemal.docx',1,'Signaturfelt og avsnitt kontrollert. Enter i avtaletekst gir nytt avsnitt med stabilt innrykk. Eksempeltekstens manuelt skrevne 5.1 er videreført fra designgrunnlaget.'),
('09_Sjekkliste_arsregnskap.docx',1,'Femkolonnetabeller med fete overskrifter, møtefelt og sideskift kontrollert. OK/I-A er tomme tabellceller, ikke avkrysningskontroller.'),
('10_Utviklingssamtale.docx',1,'Tre hovedgrupper og undernummerering 1.1/1.2 osv. Kontrollert for overlapp og plass til utfyllingsfelt.'),
('11_Diagrammer.docx',1,'Begge diagrammer, farger og bildetekster kontrollert. Datoredigering ettertestet med masterens søylediagram; se funksjonstest.'),
('12_Tekststiler_Clan.docx',1,'Typografi og lister kontrollert. Enter gir nytt punkt 3; Enter på tomt punkt avslutter listen og gir Normal.'),
('13_Tekststiler_Arial.docx',1,'Arial-varianten, tabell, signaturer og summer kontrollert.'),
('14_Marger_og_tekstflyt.docx',1,'Lang H1 brytes over to linjer. Tekstflyt og marger kontrollert.'),
('15_Avstand_mellom_tekststiler.docx',1,'Avstander rundt overskrifter og avsnitt kontrollert.'),
('16_Avstand_tabeller_og_signatur.docx',1,'Tabelltyper, avstander, fete hoder og signaturer kontrollert uten overlapp.'),
('17_Flernivanummerering_og_saksliste.docx',1,'1.1/1.2/1.3/1.3.1 og Sak 1–3 kontrollert. Enter fortsetter til 1.3.2; Reduser innrykk gir 1.4.')]
examples=[{'file':f,'clean_pages':p,'result':'kontrollert','observation':note.replace('Datoredigering','Dataredigering'),'test_copy':f'Word-QA-2026-10-21/ettertest-v3/{f}','delivery_sha256':sha256((D/f).read_bytes()).hexdigest()} for f,p,note in examples]
checks=[
('Opprettelse og lagring','DOTX åpnet som nytt Dokument. Utfylt H1 ble foreslått som navn ved første lagring til DOCX.','testkopier/QA Kontroll av Virke master.docx'),
('Stiler og tegnavstand','Kort anbefalt stiloversikt med Normal og H1–H3. H1/H2 viste Normal/0 pt i avansert skriftdialog. Alle stildefinisjoner beholdt.',None),
('Avtale og signatur','Ny Avtale-blokk satt inn. Rosa signaturfelt kan fylles ut. Brødtekst blir vanlig svart tekst; Enter lager nytt avsnitt.','native-bevis/QA Avtale Enter bestatt.docx'),
('Tabellhoder, rad og kolonne','Første/andre og ny overskriftscelle beholdt fet 7,5 pt etter skriving. Tab fra siste celle laget ny rad med normal 7,5 pt tekst.','testkopier/QA Tabell fet ny kolonne ny rad bestatt.docx'),
('Lovtekst','Innsetting beholder alle tre nummereringsnivåer: 1., 1.1 og 1.1.1.','testkopier/QA Lovtekst tre nivaer bestatt.docx'),
('Liggende tabell','Ni kolonner passet med liggende A4 og eksplisitte marger 1,25 cm venstre/høyre, 3,52 cm topp, 2,49 cm bunn.','testkopier/QA Liggende tabell bestatt.docx'),
('Feltflyt','Tab og Skift+Tab i Utfyllingsmal flyttet mellom felt. Utfylt tekst ble svart; faste Navn-etiketter ble stående.',None),
('Innholdsfortegnelse og sidetall','Tre nivåer og nytt kapittel på side 3 tatt med ved full oppdatering. Rosa punkter/innrykk og grå undernivåer bestod. H1-sidetall blir svarte.','testkopier/QA TOC tre nivaer fulloppdatert.docx'),
('Topp- og bunntekst','Begge blokker satt inn og fylt ut. PAGE/NUMPAGES reagerte på sideskift; Side 1 av 2/3 kontrollert.','testkopier/QA TOC tre nivaer fulloppdatert.docx'),
('Diagramdata','Søylediagram satt inn og dokumentet lagret. Rediger data i Excel åpnet datasettet; B2 endret fra 4,3 til 8, og stolpen i Word fulgte endringen.','testkopier/QA Diagram blokk.docx'),
('Veiledning','Alle sju DOCX-sider gjennomgått i Word. PDF eksportert fra Word og alle sju PDF-sider inspisert. Sidetall 1–7 av 7, fonter innebygd og tre kildelenker bevart.','guide-review.json')]
checks=[{'name':n,'result':'bestått med dokumenterte forbehold' if n=='Innholdsfortegnelse og sidetall' else 'bestått','observation':o,'evidence':str(QA/e) if e else 'Interaktiv Word-observasjon i denne chatten'} for n,o,e in checks]
blocks=json.loads((QA/'strukturkontroll.json').read_text())[-1]['autotext_blocks']
limitations=[
'Kontrollert i Microsoft Word for Mac. Windows og Word på nett er ikke testet; Virke-fanen må kontrolleres i Virkes Windows-miljø før utrulling der.',
'Innholdsfortegnelsens øverste nivå får svarte sidetall ved full oppdatering. Grå sidetall på dette nivået krever en manuell sluttjustering før PDF dersom Figma skal følges helt. Den øverste skillelinjen er også et gjenværende visuelt avvik fra Figma-referansen.',
'Blokkene med ni kolonner endrer ikke sideorientering eller marger automatisk. Bruk eksempel 04/05 eller sett dette eksplisitt som beskrevet i veiledningen.',
'Lagre dokumentet etter innsetting av et diagram, før Rediger data i Excel. Forsøk før første lagring gav feilmelding; etter lagring virket redigeringen.',
'Vanlig skriving i rosa plassholdere blir svart. Innliming kan beholde rosa formatering; kontroller skriftfargen etter innliming.',
'Eksemplene har fortsatt relativ tilknytning til masteren, med automatisk stiloppdatering avslått. Ferdige Styres-bruksmaler skal distribueres som selvstendige DOCX-filer. De 40 konverteringene inngår ikke i denne iterasjonen.',
'Avtaleeksempelets manuelt skrevne 5.1 og øvrig eksempelinnhold er videreført fra designgrunnlaget. Automatisk flernivånummerering er testet særskilt i Lovtekst og eksempel 17.'
]
report={'date':'2026-10-21','application':'Microsoft Word for Mac','version':word_version,'method':'Codex innebygd datamaskinbruk. Testendringer kun i kopier. Supplerende LibreOffice-render er separat og erstatter ikke Word-testene.','scope':'Master, 17 eksempler, 21 AutoText-blokker og veiledning 1.2.','examples':examples,'functional_checks':checks,'autotext_insertion':{'all_inserted':blocks,'count':len(blocks),'result':'Alle 21 satt inn i Word. Avtale, tabellhoder og Lovtekst ettertestet etter retting i siste masterrevisjon.'},'limitations':limitations,'test_artifacts':'QA-kopier kan inneholde testtekst og skal ikke distribueres.'}
js(QA/'native-results.json',report)
lines=['Virke – testresultater 21. oktober 2026',f'Microsoft Word for Mac {word_version}','',report['method'],'','OMFANG',report['scope'],'','FUNKSJONER']
for c in checks:lines.extend([f"• {c['name']}: {c['observation']}"])
lines.extend(['','ALLE EKSEMPLER'])
for e in examples:lines.append(f"{e['file']} ({e['clean_pages']} side(r)): {e['observation']}")
lines.extend(['','21 BLOKKER SATT INN',', '.join(blocks),'','KJENTE FORBEHOLD'])
lines.extend('• '+s for s in limitations)
lines.extend(['','VERN AV ORIGINALER','Innovakk-originalen fra Hans’ e-post og leveransen fra 16. oktober er bevart uendret. Alle 425 stildefinisjoner fra forrige leveranse er beholdt. Diagrammer, innebygde regneark, bilder og nummereringsdeler er bevart byte for byte. Bare nødvendig innhold, formatering, innstillinger og båndknapper er revidert.','','REN LEVERANSE','Ingen testkopier eller QA-tekst følger denne pakken. Veiledningen er redigerbar DOCX; PDF er eksportert fra samme fil i Word.'])
write(D/'TESTRESULTATER.txt','\n'.join(lines)+'\n');write(QA/'Testrapport.txt','\n'.join(lines)+'\n')
write(D/'START_HER.txt','''Virke – revidert master og eksempler
21. oktober 2026

Denne pakken inneholder DOTX-master, 17 designeksempler, fonter og
brukerveiledning 1.2 i redigerbar Word-versjon og PDF.

1. Installer begge fontene i Fonter og start Word på nytt.
2. Les Virke_mastermal_brukerveiledning.docx eller PDF-utgaven.
3. Åpne Virke_master_test.dotx fra Finder/Filutforsker for et nytt dokument.
   Fyll ut tittelen og lagre som en egen DOCX-fil.
4. På Mac: Virke > Virke-blokker. Velg Virke_master_test.dotx under Søk i,
   marker en blokk og trykk Sett inn. Bruk et tomt avsnitt som innsettingspunkt.
5. Lag en kopi når du vil fylle ut et av eksempeldokumentene.

Masteren heter fortsatt «test» i denne detaljeringen. De 40 gamle
Styres-malene skal konverteres i neste fase.

Eksemplene har relativ tilknytning til DOTX-masteren. Automatisk
stiloppdatering er avslått. Behold master og eksempler samlet når du
vil bruke masterens blokker. Ved flytting kan masteren tilknyttes på nytt.
Ferdige Styres-bruksmaler skal være selvstendige DOCX-filer i Virkes
Microsoft 365-arkiv; eksisterende utfylte dokumenter skal beholde sitt utseende.

Word for Mac er testet direkte. Windows og Word på nett er ikke verifisert.
Se TESTRESULTATER.txt for hva som er kontrollert og kjente forbehold.
''')
write(D/'ENDRINGER_OG_ETTERKONTROLL.txt','''Endringer og etterkontroll – 21. oktober 2026

• Ryddet anbefalt stiloversikt uten å slette de 425 stildefinisjonene.
• H1, H2 og tittel har 0 pt tegnavstand. Figma-størrelser og linjehøyder beholdt.
• Fete tabelloverskrifter, også første celle og etter utfylling/ny kolonne.
  Fjernet kolliderende fet-innstilling i tabellstilens betingede første rad.
• Tabellceller er toppjustert, med justert luft oppe/nede og gjentatt overskriftsrad.
• Ny Avtale-blokk med redigerbare avsnitt og rosa signaturfelt.
  Enter fungerer etter utfylling av avsnittsteksten.
• Virke-fane med Virke-blokker og Oppdater felt, uten makroer.
• Masteren inneholder 21 Autotekst-oppføringer med tydeligere navn.
• Lovtekst beholder tredje nummereringsnivå etter innsetting.
• Innholdsfortegnelsen bruker H1–H3, rosa punkter og grå undernivåer.
  Eksempel 07 har et faktisk oppdatert feltresultat fra Word.
• Nytt dokument fra DOTX foreslår utfylt H1 som navn ved første lagring.
• Brukerveiledning 1.2 er laget som DOCX med masterens egne stiler,
  tabeller, hjelpetekst og lister. PDF er eksportert fra Word.

Originalen fra Innovakk og leveransen fra 16. oktober er urørt.
Diagrammer, bilder, innebygde regneark og nummereringsdeler er bevart.
Alle 17 eksemplene og masterens 21 blokker er kontrollert i Word for Mac.
Testendringer ligger bare i QA-kopier. Se TESTRESULTATER.txt for detaljer.

Kjente forbehold: H1-sidetall i innholdsfortegnelsen blir svarte ved full
oppdatering; øverste skillelinje avviker fra Figma. Liggende tabellblokker
krever riktig sideoppsett. Diagramdata redigeres etter at dokumentet er
lagret. Innlimt tekst kan kreve at skriftfargen settes til svart.
Windows og Word på nett gjenstår som egne miljøtester.

De 40 Styres-malene er neste fase, med selvstendig DOCX-distribusjon.
''')
for ext in ('docx','pdf'):
    shutil.copy2(ROOT/f'output/{ext}/Virke_mastermal_brukerveiledning.{ext}',D/f'Virke_mastermal_brukerveiledning.{ext}')
# Preserve the old progress record, then replace it with the current result.
p=ROOT/'docs/iterasjon-2026-10-21.md';history=QA/'referanser/iterasjon-for-sluttrapport.md'
if not history.exists():shutil.copy2(p,history)
text='''# Virke Word – designerfeedback, ferdig detaljering 21. oktober 2026

Denne iterasjonen er ferdig kontrollert i Microsoft Word for Mac. Leveransen omfatter revidert DOTX, 17 rene eksempler, fonter og brukerveiledning 1.2 som DOCX og Word-eksportert PDF. De 40 Styres-konverteringene er neste fase.

## Kilder og vern

Innovakks original fra Hans’ e-post 15. oktober ligger uendret i `Word-QA-2026-10-21/referanser/Virke_mal_02_original_Innovakk.dotx`; proveniens og SHA-256 er kontrollert. 16. oktober-leveransens Word-filer er byteidentiske med sikkerhetskopien `for-iterasjon-2026-10-21.zip`. Figma-fil `lVyfrpIkkovSokMCJrG3GD`, feedback `3467:2072`; relevante TOC-, avtale-, tabell- og avstandsreferanser er gjennomgått. Historiske mellomstatuser er bevart under `referanser/iterasjon-for-sluttrapport.md`.

## Gjennomførte rettelser

Ryddet stilpanel, 0 pt tegnavstand i H1/H2/tittel, stabile fete tabellhoder, toppjusterte celler, gjentatt tabellhode, rosa signaturplassholdere, ny redigerbar Avtale-blokk, Virke-fane til Autotekst og feltoppdatering, korrekte nummernivåer i Lovtekst og H1 som foreslått førstegangsfilnavn. Alle 425 eksisterende stil-ID-er og originale diagram-/regneark-/bilde-/nummereringsdeler er bevart. 21 brukervendte blokker er tilgjengelige; tekniske plassholdere er beholdt.

Word-test avdekket og utløste ytterligere rettinger: stilpanelfilter skjulte Normal; doble betingede fet-innstillinger slo av fet skrift etter utfylling; permanente avtalefelt hindret Enter; siste Lovtekst-avsnitt mistet nivå 3 ved innsetting. Disse er rettet og ettertestet.

## Verifikasjon

`Word-QA-2026-10-21/native-results.json` inneholder filvis resultat og funksjonstester. Alle 17 eksempler er gjennomgått i Word, med testendringer bare i kopier. Alle 21 blokker er satt inn; berørte blokker er ettertestet etter siste retting. Testene omfatter utfylling, Tab/Skift+Tab, nye tabellrader/-kolonner, lang norsk tekst over fem sider, innrykk, nummerering, liggende sider, feltoppdatering, diagramdata i Excel, sidetall og første lagring.

Supplerende render av master og eksempler omfatter 19 sider; alle er visuelt inspisert etter siste endring. Veiledningens sju sider er kontrollert både som DOCX-render, i Word og i den eksporterte PDF-en. `guide-review.json` binder dette til filhashene. `verify_delivery.py` kontrollerer pakker, relasjoner, stil-ID-er, originale binærdeler og fravær av QA-tekst.

## Avklarte begrensninger

'''+''.join('- '+s+'\n' for s in limitations)+'''
## Veiledning og neste iterasjon

JSON-kilde: `docs/brukerveiledning-master.json`. DOCX bygges av `scripts/virke-build/build_master_guide_docx.py` med den bevarte masteren og komponentkontrakten i `tmp/master-guide-docx/artifact.md`. Alle 425 stiler er identiske med masteren; lister får bare nye nummereringsinstanser. PDF eksporteres manuelt fra Word med innebygd datamaskinbruk. Etter visuell kontroll registreres den med `build_master_guide.py --word-pdf <fil> --review-record <json>`; manifestet lagrer versjon, dato og master-/DOCX-/PDF-hasher.

Ved ny bygging av denne revisjonen kan `revise_designer_feedback.py --native-toc-cache Word-QA-2026-10-21/native-bevis/07-oppdatert-innholdsfortegnelse.xml` gjenbruke Word-resultatet for eksempel 07. Skriptet avviser gjenbruk hvis innhold utenfor fortegnelsen/bokmerkene har endret seg. Nye innholdsendringer krever ny fulloppdatering og kontroll i Word.

Leveransemappe: `Virke-revidert-2026-10-21`. ZIP skal bygges av denne rene mappen. Ingen QA-kopier skal følge med. Ingen commit, push eller publisering er utført.
'''
write(p,text)
print(f'Recorded {len(examples)} example checks, {len(blocks)} blocks, {len(checks)} functional checks; Word {word_version}.')
