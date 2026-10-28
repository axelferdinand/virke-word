# Virke Word

## Gjeldende leveranse og veiledning

- Brukerens avklarte omfang er én mastermal med ny design og oppdatering av 40 Styres-maler. Historiske forslag i `docs/PRD.md` om større omfang eller automatisk oppdatering ved åpning er ikke gjeldende beslutninger.
- Masteren vedlikeholdes som DOTX. Ferdige bruksmaler skal være selvstendige DOCX-filer i Virkes Microsoft 365-arkiv, uten nødvendig avhengighet til masterfilens plassering og uten automatisk stiloppdatering. Dette er besluttet leveranseretning; ikke anta at eksisterende testpakker allerede er konvertert til denne distribusjonen.
- Fontene skal følge leveransen i en egen mappe. Word på skrivebordet er hovedverktøy; ikke lov identisk funksjonalitet i Word på nett uten særskilt test.
- Brukeren har uttrykkelig bedt om at PDF-brukerveiledningen revideres ved relevante endringer i neste iterasjon og senere arbeid.
- Brukerens beslutning 21. oktober: veiledningen vedlikeholdes som redigerbar DOCX med masterens komponenter. PDF eksporteres til slutt fra Word.
- Når stiler, felt, Autotekst, tabeller, nummerering, innholdsfortegnelse, diagrammer, fonter, filnavn eller distribusjon endres: oppdater `docs/brukerveiledning-master.json`, bygg DOCX med `scripts/virke-build/build_master_guide_docx.py`, kontroller alle sider i Word og eksporter PDF. Inspiser alle rendrerte sider og registrer eksporten med `scripts/virke-build/build_master_guide.py --word-pdf <fil> --review-record <kontroll.json>`.
- Oppdater veiledningsversjon, dato og masterens SHA-256 i `docs/brukerveiledning-master-manifest.json`. Hvis en endring ikke påvirker brukerveiledningen, noter vurderingen i manifestet.
- Faste arbeidsstier er `output/docx/Virke_mastermal_brukerveiledning.docx` og `output/pdf/Virke_mastermal_brukerveiledning.pdf`. Legg begge verifiserte filer i aktuell leveransemappe og oppdater tilhørende ZIP når de inngår i leveransen.
- Ikke publiser testkopier med QA-tekst. Bevar før-versjoner. Ikke commit eller push uten brukerens uttrykkelige beskjed.
- Veiledning 2.0 er også en komponentguide: alle 21 brukervendte Autotekst-blokker vises som faktiske redigerbare komponenter. Bevar full dekning ved senere revisjoner, med levende modulhenvisninger og egne liggende sider for de brede tabellene. Designavvik fra Figma skal ikke omtales i brukerveiledningen.
- Brukerveiledningen skal beskrive praktisk bruk. Ikke ta med designavvik fra Figma eller instrukser om kosmetisk sluttjustering; slike vurderinger hører hjemme i designer-/testrapporten.
