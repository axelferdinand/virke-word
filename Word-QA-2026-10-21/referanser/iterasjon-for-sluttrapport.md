# Virke Word – designerfeedback 20. oktober 2026

Arbeidsstatus: pågår. Dette er en detaljering av masteren og eksemplene før neste fase med 40 Styres-maler. De 40 er ikke konvertert i denne iterasjonen.

## Kilder og vern av originaler

- Innovakks original: `Word-QA-2026-10-21/referanser/Virke_mal_02_original_Innovakk.dotx`, hentet uendret fra Hans' e-post 15.10.2026. Proveniens og SHA-256 ligger ved.
- Forrige leveranse er bevart uendret i `Virke-alle-Figma-eksempler-2026-10-16` og som `Word-QA-2026-10-21/for-iterasjon-2026-10-21.zip`.
- Ny arbeidsleveranse: `Virke-revidert-2026-10-21`. Interaktive testendringer skal bare lagres under QA-mappen.
- Figma: `lVyfrpIkkovSokMCJrG3GD`, Word-side `2631:9`, feedbackramme `3467:2072`. Alle kommentarinstanser lest 21.10.2026.
- Referanser lest med designkontekst og skjermbilde: innholdsfortegnelse `3130:3743`, avtale/TOC `3338:6598`, tabellhode `2804:80745`, avstander `2652:206142`.

## Krav og nødvendig sluttbevis

Alle punkter er åpne inntil implementasjon og relevante Word-tester er dokumentert.

1. Rydd stiler: et kort, sortert sett med relevante stiler i galleri/panel, uten å slette stildefinisjoner som innhold, nummerering og blokker trenger.
2. H1/H2: 0 pt tegnavstand, også koblede tegnstiler og tittelstil. Bevar riktige størrelser og linjehøyder; Figma angir H1 18 pt, H2 13,5 pt, brødtekst 9 pt, tabelltekst 7,5 pt.
3. Tabellhode: alle celler, inkludert første, er fete før og etter utfylling, samt etter ny rad/kolonne der relevant.
4. Tabellceller: toppjustering og balansert topp-/bunnmarg. Rader vokser med innhold, lange norske verdier klippes ikke, overskriftsrad gjentas.
5. Signatur: rosa redigerbare navn-/tittelfelt i avtaleeksemplene og gjenbruksblokken. Utfylling gir svart tekst, labels og tabellstruktur består.
6. Innsetting: en synlig og testet vei fra verktøylinjen til Virke-blokkene. Signatur, avtale og innholdsfortegnelse skal være lette å finne. Ikke fjern brukerens eller Words generelle byggeblokker for å rydde en dokumentmal.
7. Avtaleblokk: dagens master har ingen oppføring med dette navnet; undersøk og legg til et funksjonelt utgangspunkt med eksisterende elementer. Bevar fri redigering, Enter og stabile innrykk.
8. Innholdsfortegnelse: H1 øverst, H2/H3 som undernivåer; ingen versaler; sekundær grå på undernivåer/sidetall; rosa punkter og innrykk fra Figma; verifiser full oppdatering etter nye kapitler og sideskift.
9. Opprettelse/lagring: åpning av DOTX skal lage nytt DOCX uten å endre masteren; første lagring skal foreslå utfylt H1 som filnavn. Test faktisk dialog, ikke bare XML. Originalens core title er tom; forrige leveranse har fast title «Virke master test».
10. PDF-veiledning: fjern versale overtitler, gi brødtekst mer linjeluft og tydeligere avstand før underoverskrifter (Emil foreslår 22,5 pt). Oppdater innsetting/lagring og dokumenter Mac/Windows-forskjeller. Versjoner og inspiser alle sider.
11. Regresjon: test feltflyt, lister/innrykk, liggende tabeller, nye rader, signaturer, diagramdata, sidetall, lagring/gjenåpning og innsetting av alle blokktyper på kopier i Word.
12. Levering: oppdatert master, eksempler, fonter, veiledning, endrings-/testoversikt og ZIP. Ingen QA-tekst i leveransefiler. Originalens byggesteiner og tekniske funksjon skal være bevart eller eksplisitt forbedret.

## Avklaringer fra kildeinspeksjon

- Originalen er makrofri DOTX med 11 brukervendte byggesteiner og tekniske plassholdere. 16. oktober-masteren har 20 Autotekst-oppføringer, 203 tekniske plassholdere og 425 stiler. Tekniske plassholdere må ikke slettes ukritisk.
- Originalen lagrer Signatur m.fl. i Hurtigdeler og Innhold_Virke i TOC-galleriet. 16. oktober-versjonen samler dem i Autotekst for Mac. En synlig fane kan bruke Words innebygde AutoTextGallery uten makroer; må prøves før valg.
- Den store tittelen i eksempel 07 bruker Tittel, som ikke inngår i TOC; brødtekstens kapittel bruker Overskrift 2. Det forklarer manglende førstelinje/feil hierarki i eksemplet, selv om feltkoden inkluderer overskriftsnivå 1–3.
- Tittel har fortsatt tegnavstand −0,5 pt. H1/H2 mangler eksplisitt nullverdi i den korrigerte masteren. Begge må kontrolleres i Word etter retting.
- Mergefield 1/2 inneholder utfyllbare møteopplysninger, ikke MERGEFIELD-felt. Den gamle PDF-beskrivelsen er misvisende og skal rettes.

## Status ved videreføring

Kildeinnhenting og uendrede sikkerhetskopier er ferdige. Skriptet `scripts/virke-build/revise_designer_feedback.py` har revidert alle 18 Word-pakker i arbeidsleveransen. `Word-QA-2026-10-21/revisjon.json` og `strukturkontroll.json` dokumenterer delendringer og strukturkontroller. Diagrammer, regneark, bilder og nummereringsdeler er bevart byte for byte.

Native Word-prototyper er testet: tom core title gir H1 som forslag til filnavn ved første lagring. DOTX åpnet som nytt Dokument1. AutoTextGallery viste bare generelle standardblokker, mens `AutoManager` åpner Autotekst-dialogen og viser masterens 20 blokker når masteren velges under Søk i. Signatur ble satt inn. Derfor bruker revisjonen AutoManager-knappen på Virke-fanen, uten makroer. En deaktivert ApplyStylesPane-knapp er utelatt.

Den reviderte masteren er ennå ikke åpnet i Word. Datamaskinbruk meldte at Mac-en er låst, og asynkront spørsmål om manuell opplåsing er sendt. Full Word-regresjon, TOC-nivåer/farger etter oppdatering, stilpanel og innsetting av alle 21 blokker gjenstår. Ikke marker noe av dette som bestått fra XML alene.

PDF 1.1 er bygget i sju sider, alle PNG-er visuelt inspisert. Manifeststatus skiller PDF-visuell kontroll fra Word-regresjon som gjenstår. PDF-en ligger også i arbeidsleveransen. Guide og manifest må bygges opp igjen hvis master eller brukerflyt endres under etterkontroll.

Supplerende render_docx-kontroll bruker LibreOffice og erstatter ikke brukerens pålagte native Word-test. Første kjøring manglet fontconfig og substituerte Times; den er ikke gyldig layoutbevis. Andre kjøring bruker `Word-QA-2026-10-21/fonts.conf` med Clan og systemfonter. Alle 19 sider for 18 filer i `render/` er nå visuelt inspisert med riktige fonter, uten synlig klipping eller overlapp. Eksempel 07 har fortsatt gammel lagret TOC-visning og må fulloppdateres i Word. `render-manifest.json` skiller supplerende kontroll fra native Word-bevis. Testkopier av de reviderte 18 filene er klargjort i `Word-QA-2026-10-21/testkopier`.

Ingen endelig ZIP er bygget, og START_HER/ENDRINGER er tydelig merket arbeidsversjon. Ikke merk målet fullført før alle krav har konkret sluttbevis.

## Brukerens oppdaterte beslutning 21. oktober

Brukerveiledningen skal vedlikeholdes som DOCX med masterens komponenter. PDF eksporteres til slutt fra den kontrollerte Word-veiledningen. Oppdater også JSON-kilden og manifestet.

## Native Word-funn etter opplåsing

Revidert DOTX opprettet Dokument3/4; utfylt H1 gav «QA Kontroll av Virke master.docx» som første lagringsforslag. H1/H2 viser Normal tegnavstand i Words avanserte skriftdialog. Virke-knappen åpner Autotekst; filteret viser 21 oppføringer. Avtale ble satt inn med rosa signaturer som kan fylles ut. Funn: det eksplisitte stilpanelfilteret skjulte Normal (rettet ved å følge Words standard Anbefalt). Nye avtaletekstfelt blokkerte Enter (rettet til midlertidige kontroller som blir vanlig tekst ved første utfylling). Ettertest pågår.
