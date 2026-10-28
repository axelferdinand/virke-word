# Virke Word – designerfeedback, ferdig detaljering 21. oktober 2026

Denne iterasjonen er ferdig kontrollert i Microsoft Word for Mac. Leveransen omfatter revidert DOTX, 17 rene eksempler, fonter og brukerveiledning 1.3 som DOCX og Word-eksportert PDF. De 40 Styres-konverteringene er neste fase.

## Kilder og vern

Innovakks original fra Hans’ e-post 15. oktober ligger uendret i `Word-QA-2026-10-21/referanser/Virke_mal_02_original_Innovakk.dotx`; proveniens og SHA-256 er kontrollert. 16. oktober-leveransens Word-filer er byteidentiske med sikkerhetskopien `for-iterasjon-2026-10-21.zip`. Figma-fil `lVyfrpIkkovSokMCJrG3GD`, feedback `3467:2072`; relevante TOC-, avtale-, tabell- og avstandsreferanser er gjennomgått. Historiske mellomstatuser er bevart under `referanser/iterasjon-for-sluttrapport.md`.

## Gjennomførte rettelser

Ryddet stilpanel, 0 pt tegnavstand i H1/H2/tittel, stabile fete tabellhoder, toppjusterte celler, gjentatt tabellhode, rosa signaturplassholdere, ny redigerbar Avtale-blokk, Virke-fane til Autotekst og feltoppdatering, korrekte nummernivåer i Lovtekst og H1 som foreslått førstegangsfilnavn. Alle 425 eksisterende stil-ID-er og originale diagram-/regneark-/bilde-/nummereringsdeler er bevart. 21 brukervendte blokker er tilgjengelige; tekniske plassholdere er beholdt.

Word-test avdekket og utløste ytterligere rettinger: stilpanelfilter skjulte Normal; doble betingede fet-innstillinger slo av fet skrift etter utfylling; permanente avtalefelt hindret Enter; siste Lovtekst-avsnitt mistet nivå 3 ved innsetting. Disse er rettet og ettertestet.

## Verifikasjon

`Word-QA-2026-10-21/native-results.json` inneholder filvis resultat og funksjonstester. Alle 17 eksempler er gjennomgått i Word, med testendringer bare i kopier. Alle 21 blokker er satt inn; berørte blokker er ettertestet etter siste retting. Testene omfatter utfylling, Tab/Skift+Tab, nye tabellrader/-kolonner, lang norsk tekst over fem sider, innrykk, nummerering, liggende sider, feltoppdatering, diagramdata i Excel, sidetall og første lagring.

Supplerende render av master og eksempler omfatter 19 sider; alle er visuelt inspisert etter siste endring. Veiledningens sju sider er kontrollert både som DOCX-render, i Word og i den eksporterte PDF-en. `guide-review.json` binder dette til filhashene. `verify_delivery.py` kontrollerer pakker, relasjoner, stil-ID-er, originale binærdeler og fravær av QA-tekst.

## Avklarte begrensninger

- Kontrollert i Microsoft Word for Mac. Windows og Word på nett er ikke testet; Virke-fanen må kontrolleres i Virkes Windows-miljø før utrulling der.
- Innholdsfortegnelsens øverste nivå får svarte sidetall ved full oppdatering, og en ekstra øverste skillelinje består. Fire isolerte Word-eksperimenter gav ingen stabil retting som samtidig bevarer vanlige overskriftsstiler, klikkbare kapitteltekster og skillelinjer mellom kapittelgruppene. Dette er gjenværende designavvik; funksjonen med innhold og sidetall er bevart.
- Blokkene med ni kolonner endrer ikke sideorientering eller marger automatisk. Bruk eksempel 04/05 eller sett dette eksplisitt som beskrevet i veiledningen.
- Lagre dokumentet etter innsetting av et diagram, før Rediger data i Excel. Forsøk før første lagring gav feilmelding; etter lagring virket redigeringen.
- Vanlig skriving i rosa plassholdere blir svart. Innliming kan beholde rosa formatering; kontroller skriftfargen etter innliming.
- Eksemplene har fortsatt relativ tilknytning til masteren, med automatisk stiloppdatering avslått. Ferdige Styres-bruksmaler skal distribueres som selvstendige DOCX-filer. De 40 konverteringene inngår ikke i denne iterasjonen.
- Avtaleeksempelets manuelt skrevne 5.1 og øvrig eksempelinnhold er videreført fra designgrunnlaget. Automatisk flernivånummerering er testet særskilt i Lovtekst og eksempel 17.

## Veiledning og neste iterasjon

JSON-kilde: `docs/brukerveiledning-master.json`. DOCX bygges av `scripts/virke-build/build_master_guide_docx.py` med den bevarte masteren og komponentkontrakten i `tmp/master-guide-docx/artifact.md`. Alle 425 stiler er identiske med masteren; lister får bare nye nummereringsinstanser. PDF eksporteres manuelt fra Word med innebygd datamaskinbruk. Etter visuell kontroll registreres den med `build_master_guide.py --word-pdf <fil> --review-record <json>`; manifestet lagrer versjon, dato og master-/DOCX-/PDF-hasher.

Ved ny bygging av denne revisjonen kan `revise_designer_feedback.py --native-toc-cache Word-QA-2026-10-21/native-bevis/07-oppdatert-innholdsfortegnelse.xml` gjenbruke Word-resultatet for eksempel 07. Skriptet avviser gjenbruk hvis innhold utenfor fortegnelsen/bokmerkene har endret seg. Nye innholdsendringer krever ny fulloppdatering og kontroll i Word.

Leveransemappe: `Virke-revidert-2026-10-21`. ZIP skal bygges av denne rene mappen. Ingen QA-kopier skal følge med. Ingen commit, push eller publisering er utført.

## Etterkontroll etter forespørsel om TOC-retting

Fire isolerte varianter ble generert fullt, lagret og visuelt kontrollert i Word for Mac. Direkte fjerning av første topplinje overlevde ikke regenerering. Grå TOC1 gjorde også kapittelteksten grå. Uten hyperlenkebryteren kunne særformatert svart kildetekst gi riktig farge, men vanlige nye H1 feilet og kapitteltekstens hyperlenkeinnpakning forsvant. Mellomlinjer alene fjernet ønskede kapittelskiller; bunnlinjer gav streker ved undernivåene. Ingen variant ble innført. Detaljer og filhasher: `Word-QA-2026-10-21/toc-etterkontroll/resultater.json`.

Master og eksempler er byteidentiske med for-versjonen av 21. oktober-pakken, bevart i `toc-etterkontroll/for-versjon`. Veiledning 1.3 fjerner designmerknaden og rådet om manuell fargejustering etter brukerens avklaring. Alle sju sider er kontrollert på nytt i Word, DOCX-render og Word-eksportert PDF. Designavvikene omtales bare i designer-/testrapporten.


## Ny leveranse med komponentguide 2.0

På brukerens forespørsel er Word-guiden utvidet til 18 sider som viser alle 21 brukervendte Autotekst-blokker som faktiske redigerbare komponenter. Tabellene, hjelpeteksten og avtaleblokken forklarer sin egen bruk. Side 2 bruker en trekolonnetabell med 21 levende PAGEREF-henvisninger. Innholdsfortegnelsen er avgrenset til stilprøven på side 3; begge ni-kolonnetabellene har liggende A4 på side 10–11. Masteren og de 17 eksempelfilene er byteidentiske med forrige leveranse.

Guidekilden har versjon 2.0. Bygg DOCX med `build_master_guide_docx.py`, oppdater alle felt i en Word-testkopi, og importer den verifiserte TOC-cachen med `refresh_master_guide_fields.py --native-docx <testkopi>`. Dette beholder masterens stiler, Autotekst, diagrammer og regneark urørt. Nye innholdsendringer krever ny feltoppdatering og visuell kontroll; ikke gjenbruk gamle feltresultater ukritisk. Eksporter deretter fra en ren kopi i Word og sluttfør med `build_master_guide.py`.

Kontroll: 18 sider i native Word-PDF og supplerende DOCX-render inspisert. Representative sider kontrollert direkte i Word; F9, tre TOC-nivåer, henvisninger og navnefelt prøvd. Klikk fra oversikten hoppet til side 10; Tab/Skift+Tab og svart utfyllingstekst fungerte. PDF har 24 interne lenker. Testredigering ble forkastet og finnes ikke i leveransen. `Word-QA-2026-10-21/komponentguide/guide-review.json` binder kontrollen til filhashene. Windows og Word på nett er fortsatt ikke testet.

Gjeldende nye leveranse er `Virke-med-komponentguide-2026-10-21.zip`, bygget fra mappen med samme navn. Forrige ZIP og guide 1.3 er bevart. Ingen commit eller push.
