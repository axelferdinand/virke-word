# Virke brukerveiledning 1.3 – dokumentkontrakt

Referanse: /Users/afg/Nettsteder/virke-word/Virke-revidert-2026-10-21/Virke_master_test.dotx
SHA-256: 74d641efa36287eeb70d34cd38036d2c09de31288e5d1d1a6e22399752e44740
Bevart kopi: reference.dotx. reference.docx er samme pakke med dokument-content-type for auditverktøy. Én side, én stående A4-inndeling. Referanse er kontrollert i native Word med alle 21 blokker og supplerende renderer i Word-QA-2026-10-21/render/Virke_master_test.

## Sidesystem
11906 × 16838 twips, topp 1995, bunn 1410, venstre/høyre 1288, topp-/bunntekst 705. Én kolonne. Behold sectPr og relasjoner. Veiledningen har sju tematiske sider med sideskift før hver H1 etter den første. Samme topp- og bunntekst på alle sider, levende PAGE og NUMPAGES.

## Stiler og komponenter
Behold styles.xml, theme, fonttabell og alle 425 definisjoner uendret. Bruk Overskrift1 18 pt, linje 1,35, etter 22,5 pt. Underoverskrifter bruker Overskrift3 10,5 pt, minst 16,8 pt, før 22,5 pt, etter 9 pt. Normal 9 pt, minst 14,4 pt. Listeavsnitt 9 pt, minst 14,4 pt, etter 6 pt, hengende innrykk 340 twips. Bildetekst 7,5 pt, minst 12 pt for sekundærtekst. Clan News/Bold fra tema. Vanlig bokstavbruk uten overtittel.

Tabell: klon Autotekst Tabell med 2 kolonner, bredde 9330 twips, tilpass to kolonnebredder fra JSON. Behold tabellstil Virke1, grå EDEDED hoder, CACACA kantlinjer 0,75 pt, toppjustering, topp 150/bunn 90 twips, sidepadding 120. Klon kroppsrader etter behov, gjenta overskriftsrad. Bruk VirkeTabelloverskrift / VirkeTabelltekst uten direkte fonts. Hjelpetekst: klon faktisk éncellet Hjelpetekst-blokk, rosa kant og VirkeHjelpetekst, padding 150, fyll med veiledning.

Lister: bruk eksisterende numId 21 for punktlister. Ny nummereringsinstans per stegsekvens klones fra numId 28 / abstractNumId 16 med startOverride 1. Eksisterende nummerdefinisjoner bevares semantisk uendret; kun nye instanser legges til.

## Innhold og slot map
JSON-kilden docs/brukerveiledning-master.json gir sju sider: start, stiler/felt, innsetting, blokkoversikt, tabeller/TOC, diagram/deling, vedlikehold. Dokumentets tre startavsnitt er innholdssloter og erstattes av disse sidene. H1/normal/H3 kopieres som avsnittsmønstre fra eksisterende stilsystem. Ingen innhold fra originale avtaleeksempler følger med. Nytt innhold i klonede tabeller er redigerbar statisk veiledning, uten utfyllingsplassholdere. Header1 SDT virke_170000716 og footer1 SDT virke_170000718 fylles gjennom tekstnoder, placeholderflagget fjernes og eksisterende grå sluttfarge brukes. Feltene PAGE/NUMPAGES beholdes. Kilder blir fungerende eksterne hyperlenker.

## Pakkevern og tillatte endringer
Inventory: package-inventory.json. Redigerbare deler: word/document.xml, word/header1.xml, word/footer1.xml, word/numbering.xml (bare nye num), word/_rels/document.xml.rels (bare nye kildehyperlenker), docProps/core.xml, docProps/app.xml, [Content_Types].xml. Alle andre deler skal være byteidentiske, inkludert glossary, customXml, stiler, bilder, diagrammer, innebygde Excel-ark og deres relasjoner. Ingen makroer. Fil lagres som DOCX.

## Verifikasjon
Audit seksjoner, stiler, felt og kontroller før/etter. Behold referansens hash. Render alle sider med render_docx og inspiser alle PNG-er. Visuell fidelity sammenligning med masteren: geometri, skrifter, topp/bunn, tabell-/hjelpetekstkomponenter. Åpne i Word, kontroller alle sju sider, oppdater felt og eksportér PDF fra Word. Endelig PDF skal gi korrekte sider 1–7 av 7 og ingen klipping. Native Word kan normalisere XML ved lagring; bevar OOXML-kilde og importer kun nødvendige feltresultater ved behov, ingen LibreOffice-resave.

Revisjon 1.3: Designmerknaden på side 5 er fjernet. Praktiske steg for full oppdatering beholdes. Masterreferansen er uendret.

## Komponentguide 2.0

Brukerens nye bestilling er en visuell, redigerbar guide som demonstrerer alle 21 offentlige Autotekst-blokker. Bevar 1.3 i Word-QA-2026-10-21/komponentguide/for-versjon. Ny leveransemappe er Virke-med-komponentguide-2026-10-21; masteren er byteidentisk med referansen over.

Planlagt sidesystem: 18 tematiske sider. Side 10 og 11 kloner liggende A4-geometri fra eksempel 04: 16838 × 11906 twips, venstre/høyre 709, topp 1995, bunn 1410. Stående sider bruker opprinnelig geometri. To inndelingsskift. Ny liggende bunntekst kloner masterens bunntekst og flytter høyre tabulator til 15420. PAGE/NUMPAGES løper gjennom alle inndelinger.

Alle offentlige docPartBody-blokker klones til hoveddokumentet. Verifiserte tekstnoder i vanlige SDT-er fylles i stedet for at feltene bygges om. Tabellvarianter beholder stil, grid, cellemarger og rader; tekst og antall eksemplariske rader tilpasses. Komponentoversikten bruker Tabell med 3 kolonner med tilpasset bredde 4050/4480/800. Style-prøver viser H1–H3 og Normal; guidekapitler bruker masterens Tittel-stil. Små veiledningsetiketter er fet Normal. Alle 425 stiler bevares uendret.

Innholdsfortegnelsens felt avgrenses med bokmerket GuideStyleSample til de tre faktiske stilprøvene. Levende PAGEREF-felt i oversiktstabellen peker til de demonstrerte modulene. Modulenes egne kontroller, tabeller, diagrammer og innebygde regneark beholdes. SDT-, tegning- og bokmerkeidentiteter gjøres unike. Header/footer-blokkene vises også inne i dokumentet som redigerbare prøver. Nye bokmerker og nummereringsinstanser er additive.

Tillatte endringer i guidepakken: document.xml, header1.xml, footer1.xml, numbering.xml (additivt), document.xml.rels (additivt), core.xml og content types. Ny footerGuideLandscape.xml for liggende sider. Glossary, theme, styles, chart XML, chart styles, chart colors og innebygde arbeidsbøker beholdes byte for byte. Native Word-feltresultater importeres målrettet etter kontroll. Masteren og de 17 separate eksemplene endres ikke.

Verifiser 21 av 21 komponenter, alle kontroller og relasjoner, levende TOC/PAGEREF/PAGE/NUMPAGES, alle sider i supplerende render, alle sider i Word og alle sider i Word-eksportert PDF. Ikke publiser QA-tekst eller omtale av Figma-avvik i guiden.


### Kontrollert sluttresultat 2.0

18 sider, alle 21 byggeblokker, 21 modulhenvisninger og tre TOC-lenker. Liggende bunntekst tømmer den arvede tabulatoren på 9330 twips og setter høyre tabulator på 15420. TOC-feltets slutt ligger i eget INNH1-avsnitt for å unngå en tom punktlinje. Native feltresultat importeres separat; 29 stil-/tema-/glossary-/diagram-/regneark-/bildedeler er byteidentiske med masteren. Alle sider i begge rendringer inspisert; Word-PDF er autoritativ.
