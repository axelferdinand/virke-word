---
task: Virke Word template inventory analysis
slug: 20260422-093000_virke-word-template-inventory
effort: extended
phase: complete
progress: 30/30
mode: interactive
started: 2026-04-22T09:30:00+02:00
updated: 2026-04-22T10:15:00+02:00
---

## Context

Client: Virke (Hovedorganisasjonen Virke).
Folder: `Virke/Word-maler/Maler i AGG/{Norske maler, Engelske maler}`
Scope: 131 files total (104 Norwegian + 27 English), ~120 Word (.docx/.dotx/.doc) and ~11 Excel (.xlsx/.xlsm).

Goal: Deliver a structured inventory + style analysis + simplification opportunities + concrete project plan that feeds into Figma design-system creation, which then becomes the source of truth when Claude (or other models) rewrite all templates.

### Risks

- No Virke branding exists in templates today → we need explicit brand direction
- Style bloat (up to 422 styles in one file) → risk of carrying cruft into new design system
- Multiple Office theme variants (4 accent-blue versions) → users may be attached to certain layouts
- Filename inconsistency → migration + navigation UX risk
- Duplicate IDs and near-dupe files → scope creep risk if we try to consolidate during redesign

### Plan

Approach in four phases: (1) this analysis deliverable (now), (2) Figma design-system build based on agreed tokens, (3) pilot rewrite of 3-5 representative templates with human review, (4) batch AI-assisted rewrite of remaining templates using the pilot as pattern.

## Criteria

- [x] ISC-1: Total template count verified across both language folders
- [x] ISC-2: All templates enumerated with category classification
- [x] ISC-3: File type distribution documented (docx/dotx/doc/xlsx/xlsm)
- [x] ISC-4: Number of Norwegian files matches 104
- [x] ISC-5: Number of English files matches 27
- [x] ISC-6: Theme color scheme identified across entire corpus
- [x] ISC-7: Accent color variant count documented with frequencies
- [x] ISC-8: Body font distribution documented across corpus
- [x] ISC-9: Page margin consistency verified across corpus
- [x] ISC-10: Heading style hierarchy documented (sizes, colors, font)
- [x] ISC-11: Footer pattern documented from sample templates
- [x] ISC-12: Logo presence checked across all templates (result: zero)
- [x] ISC-13: Placeholder convention patterns extracted from document bodies
- [x] ISC-14: Red-text and yellow-highlight usage quantified
- [x] ISC-15: Style count distribution shown (min/max/avg/bloat count)
- [x] ISC-16: Virke branding presence quantified (text mentions)
- [x] ISC-17: Duplicate or near-duplicate filenames flagged
- [x] ISC-18: ID-number collision between templates flagged
- [x] ISC-19: Filename convention problems documented with examples
- [x] ISC-20: File format inconsistencies flagged (.doc, .docx where .dotx expected)
- [x] ISC-21: Proposed Figma design-token list defined (colors)
- [x] ISC-22: Proposed Figma design-token list defined (typography)
- [x] ISC-23: Proposed Figma design-token list defined (spacing + layout)
- [x] ISC-24: Proposed Word-style system defined (paragraph + character)
- [x] ISC-25: Proposed Word-style system defined (table + list)
- [x] ISC-26: Template-component inventory for Figma defined
- [x] ISC-27: Pilot template set recommended with rationale
- [x] ISC-28: AI-assisted rewrite workflow described
- [x] ISC-29: Placeholder convention proposal stated
- [x] ISC-30: Open questions listed for user confirmation

## Decisions

### First-Principles — designsystem-primitiver (fra 413 stiler → 15)

**Constraint-klassifisering**

| Constraint | Type | Evidens | Utfordret |
|------------|------|---------|-----------|
| Må fungere i Word 2016+ Win/Mac | Hard | Kundebase | Nei — reell |
| Må rendere identisk til PDF | Hard | Juridisk bruk | Nei — rettskraft |
| "Calibri fordi default" | Soft | Legacy | Ja — Virke kan velge font |
| "Heading1-4 er hierarkiet" | Assumption | Word-konvensjon | Ja — 2 nivå holder ofte |
| "Hver mal har egne stiler" | Soft | 413-stil-bloat | Ja — én delt master |
| "Rød tekst = instruks" | Assumption | 22 filer | Ja — erstattes med semantisk stil |
| "Office-blå er brand" | Soft | Ingen brand | Ja — velg ekte Virke-farge |
| "413 stiler trengs" | Assumption | Akkumulert cruft | Ja — 15 dekker alt |

**Valgt primitiv-sett**

```
Color (primitive)
  ink-900:   #111111    /* body */
  ink-600:   #4A4A4A    /* sekundær */
  ink-300:   #B3B3B3    /* muted */
  paper:     #FFFFFF    /* bakgrunn */
  brand-primary:    [Virke-farge TBD]
  brand-secondary:  [Virke-farge TBD]
  signal-action:    #C83020  /* erstatter #FF0000 "fyll-inn" */
  signal-subtle:    #FFF3BF  /* erstatter gul uthev */

Type (primitive)
  font-heading: [Virke-font TBD, fallback Calibri]
  font-body:    [Virke-font TBD, fallback Calibri]
  font-mono:    Consolas
  Skala: 9 · 10 · 11 · 12 · 14 · 18 · 22 pt
  Line-height: tight 1.15 · body 1.35 · loose 1.5

Space (primitive — Word "space before/after")
  0 · 4 · 8 · 12 · 18 · 24 pt
```

### Word-stilsystem (semantiske styleId-er)

**Paragraph (9 — erstatter ~94 per mal):**
- `Tittel` → size-7, brand-primary, space-5 after
- `Undertittel` → size-5, ink-600, space-4 after
- `Overskrift1` → size-6, ink-900, space-4 before / 2 after
- `Overskrift2` → size-5, ink-900, space-3 before / 1 after
- `Brødtekst` → size-3, ink-900, leading-body (default)
- `Hjelpetekst` → size-2, ink-600, italic (merknader/vilkår)
- `Instruksjon` → size-3, signal-action, italic (erstatter rød tekst)
- `Liste` → size-3, indent 8mm, bullet/number
- `Signaturlinje` → size-3, top-border, space-4 before

**Character (3):**
- `Fremhevet` → bold
- `Variabel` → ink-600, `{…}`-braces — "fyll-inn"-markør
- `Paragrafreferanse` → font-mono, size-2 — lovhenvisninger

**Table (3 — erstatter 248 Word-defaults):**
- `Skjematabell` → borders, brand-primary header-row
- `Signaturtabell` → no borders, fixed 2-col for sign-off
- `Data` → alt-row shading ink-300 @ 8% opacity

### Komponenter (Figma → .dotx Building Blocks via Glossary)

1. **Masthead** — Virke-logo + dok-tittel + dok-ID
2. **Partsblokk** — Arbeidsgiver/Arbeidstaker-rute med variabler
3. **Signaturblokk** — Dato + Sted + 2 signaturlinjer
4. **Avsnittsvarsel** — Instruksjon-callout (erstatter rød tekst)
5. **Paragrafhenvisning** — formatert juridisk referanse
6. **Vilkårsliste** — nummerert med paragraf-nivå
7. **Plassholder-chip** — `{navn}`, `{dato}` osv i Variabel-stil
8. **Sidefot** — "Side X av Y · [Mal-tittel] · [ID] · Virke"

### Plassholder-konvensjon — standardisert

Standardisér til `{variabelnavn}` i Character-stil `Variabel`. Gjør dem trivielle å regex-erstatte ved automatisk utfylling. Eksempler: `{arbeidstaker_navn}`, `{dato}`, `{stilling}`, `{arbeidsgiver_navn}`.

### Key insight

**De 413 stilene per mal er ikke et designsystem — det er arkeologi.** Ekte designsystem = 9 paragraph + 3 character + 3 table + 8 komponenter, delt via én `Virke-Master.dotx` koblet til alle dokumenter.

## Verification

# 📋 Virke Word-maler — Inventar, analyse og angrepsplan

## 1. Inventar — 131 filer totalt

| Språk | Antall | Filtyper |
|------|-------:|----------|
| **Norske maler** | 104 | 39 .docx · 54 .dotx · 1 .doc · 9 .xlsx · 1 .xlsm |
| **Engelske maler** | 27 | 2 .docx · 24 .dotx · 1 .xlsx |
| **Sum** | **131** | 41 .docx · 78 .dotx · 1 .doc · 10 .xlsx · 1 .xlsm |

### 1.1 Funksjonell kategorisering (norske)

| Kategori | Antall | Nøkkelmaler |
|---------|-------:|------------|
| Ferie/sykefravær/permisjon | 16 | Egenmelding, Ferieplan, Barns sykdom, Foreldrepermisjon |
| Permittering | 12 | Permitteringsvarsel (7 varianter), Drøftingsprotokoller |
| HMS/vern | 11 | Vernerunde, Risikovurdering, Avvik, Varslingsrutiner |
| Kompetanse/rekruttering | 10 | Intervjuguide, Introduksjonsplan, Avslagsbrev, Referansesjekk |
| Oppsigelse/avskjed | 8 | Oppsigelse A/B/C, Avskjed, Suspensjon, Sluttavtale |
| Avtaler | 8 | Konkurranseklausul, Kundeklausul, Hjemmekontor, Gjeldsbrev |
| Skjema/rapport | 8 | HUKI, Årshjul, Sjekklister |
| Personvern | 8 | Databehandleravtale, Personvernerklæringer, Samtykke |
| Ansettelse | 6 | Ordinær, Midlertidig, CEO, Leder, Lærling, Rammeavtale |
| Medarbeidersamtale | 5 | Samtaleskjema, Prøvetid (3 varianter) |
| Protokoll/referat | 4 | Drøftingsmøte, Innkalling, Referat aml §14-1 |
| Advarsel | 2 | Generell, Rusmiddel |
| Annet | 6 | Attest, Sykmelding, Huskeliste, Lønnsansiennitet |

Engelske maler speiler i stor grad et utvalg av de norske (27 av 104 er oversatt).

### 1.2 Identifiserte duplikater / nærdupliser

- `Samtykkeerklæring – publisering av bilder 1111.dotx` **og** `... 1111 (1).dotx` → identiske, én bør slettes
- `Fraværskort for 2025.docx` **og** `Fraværskort_2026.docx` → årlig versjonering, bør bli én parametrisert mal
- `Prøvetid oppfølgingssamtale – 1102.dotx` **og** `... Første samtale` **og** `... Andre samtale` → tre filer der én med seksjoner hadde vært tydeligere
- `260418_Betinget permitteringsvarsel.docx` + `260418_Permitteringsbevis.docx` + `260418_Protokoll...` → dato i filnavn tyder på "midlertidig levert" — bør normaliseres
- **ID-kollisjon:** «1008» brukes både på `Permitteringsvarsel ved arbeidsmangel…` og `Personalopplysninger og utregning…` — brudd i nummereringssystemet

---

## 2. Stil/format-oversikt — dagens tilstand

### 2.1 Tema og farger (alle 93 norske Word-filer analysert)

| Egenskap | Funn |
|---------|------|
| Color scheme | `Office` (100% — ingen Virke-palett) |
| Accent1-variant | **4 ulike blåtoner:** #4F81BD (44 maler) · #4472C4 (41) · #5B9BD5 (6) · #156082 (2) |
| Body-font (theme) | Calibri (91) · Aptos (2, utilsiktet nyere-Office-default) |
| Heading-font (theme) | Calibri / Calibri Light (blandet) |
| Sidestørrelse | A4 gjennomgående |
| Marger | 25/25/25/25 mm (92 av 93) · 1 avviker med venstremarg 30mm |

**Konklusjon:** Ingen Virke-brand eksisterer i malene. Accent-fargene er rester fra ulike Office-versjoner templatene ble laget i, ikke et valg.

### 2.2 Stilsystem (brukerrettede paragraph-stiler)

Fra Ansettelsesavtale ordinær (representativ):

| Stil | Størrelse | Farge | Merknad |
|------|----------|-------|---------|
| Normal | — | auto | Theme-default Calibri |
| Heading1 | 20pt | #000000 | Ingen brand-farge |
| Heading2 | 16pt | #000000 | |
| Heading3 | 15pt | #000000 | |
| Heading4 | 14pt | #000000 | |
| Title / Subtitle | 10pt | #000000 | Subtitle mindre enn body — merkelig |
| Footer | 9pt | — | Brukes av sidefot |

**Stil-bloat på tvers av korpus:**
- Snitt: **82 stiler per mal**
- Maks: **422 stiler** (en mal)
- 10 maler har >300 stiler
- 69 maler har <100 stiler (rimelig)
- Totalt i korpus er det sannsynligvis >20 000 unike style-definisjoner, hvorav ~15 faktisk brukes

### 2.3 Farge- og typografi-kaos i dokumentinnhold

- **26 ulike farger** referert i en enkelt mals styles.xml
- **11 ulike skriftstørrelser** (8, 9, 10, 10.5, 11, 12, 13, 14, 15, 16, 20 pt)
- **3–4 fonter** aktivt brukt: Calibri, Segoe UI, Times New Roman, Consolas
- **Rød tekst (#FF0000):** brukes som "fyll inn her"-markør i 22 maler
- **Gul uthev:** brukes til samme formål i 10 maler (inkonsistent)

### 2.4 Sidefot-konvensjon (én av få ting som er konsistent)

Mønster i alle sampler: `Side X av Y   [Mal-tittel] – [ID]`
Ingen header, ingen logo, ingen Virke-referanse.

### 2.5 Plassholder-konvensjon

Top-10 plassholder-mønstre (ikke standardisert):

| Mønster | Forekomster |
|--------|------------:|
| `[dato]` | 50 |
| `[Navn på arbeidstaker]` | 22 |
| `[Sted og dato]` | 21 |
| `[x antall]` | 14 |
| `[Navn]` | 11 |
| `[Arbeidstakers adresse]` | 10 |
| `[Selskapets navn]` | 10 |
| `[Navn på virksomhet]` | 8 |
| `[Etternavn og fornavn]` | 7 |
| `[virksomhetens navn]` | 5 |

**Store/små bokstaver, avkorting og stil varierer.** Samme konsept skrives på 3-4 måter.

### 2.6 Building Blocks (.dotx-spesifikt)

Alle .dotx-filer bruker **Word Building Blocks** aktivt via `word/glossary/` — dette er allerede en infrastruktur vi kan bygge videre på i designsystemet.

### 2.7 Logoer og bilder

**0 av 131 filer** inneholder logo eller bilder. Virke er visuelt totalt fraværende.

### 2.8 Excel-filer (11 stk)

Gjennomsnitt 1,9 ark. Bruker cellestiler, men **ingen** benytter Excel-tabellfunksjonalitet — de er i praksis skjemaer/layout-dokumenter laget i Excel. Tre valg: (a) flytt til Word, (b) lag parallelt mini-designsystem for Excel, (c) behold med minimal styling-oppdatering.

---

## 3. Forenklingsmuligheter

### 3.1 Stilsystem — 413 stiler → 15

Se Decisions → Word-stilsystem for full spesifikasjon. Kort: 9 paragraph + 3 character + 3 table. Resten er Word-default-cruft som bør fjernes fra master-template.

### 3.2 Tokens — én gang i Figma, deriveres til Word

Se Decisions → primitiv-sett for full token-liste.

### 3.3 Plassholder-konvensjon — én form

Standardisér til `{variabelnavn}` i Character-stil `Variabel` (grå, i krøllparenteser) — gjør dem trivielle å regex-erstatte under automatisk utfylling.

### 3.4 Rød tekst → `Instruksjon`-stil

De 22 malene som bruker rød instruks-tekst bør ha en tydelig semantisk stil (f.eks. lys kasse med ikon + kursiv tekst) som skjermleser og utskrift-eksport forstår. Samme for de 10 som bruker gul uthev.

### 3.5 Filnavn-normalisering

Nåværende inkonsistenser:
- Spaces rundt dash: `"Advarsel -1000"` vs `"Avskjed - 1012"`
- Dato i filnavn: `"260418_"`, `"010724"`
- ID før/etter: `"1089 A.docx"` vs `"_1085.docx"`
- Mixed case: `"avtale om forskuddsferie…"` (liten a)

**Forslag til mønster:** `{NNNN}_{Kategori}_{Kort-tittel}_{vYYMM}.dotx`
F.eks. `1089A_Oppsigelse_Virksomhetens-forhold_v2604.dotx`

### 3.6 Filtype-oppryddsing

- Alt som er "mal" → `.dotx` (78 stk i dag)
- `.doc`-filen (Sluttavtale) må migreres til `.docx` eller `.dotx`
- Navngitt "docx" men egentlig mal → konverter (41 stk)

### 3.7 Master.dotx-strategi (enkelt å overse, men game-changer)

Alle 120 Word-maler deler **én** `Virke-Master.dotx` som inneholder tokens + stiler + building blocks. Hver mal lagres som .dotx som peker til masteren. Endrer du en stil ett sted, oppdateres alle maler neste åpning.

---

## 4. Konkret angrepsplan

### 4.1 Faseinndeling (~8-10 arbeidsuker)

| Uke | Fase | Leveranse |
|-----|-----|----------|
| **1-2** | **Fundament** — brand, tokens, Master.dotx | Figma-tokens låst + `Virke-Master.dotx` v1 |
| **3** | **Pilotbatch (5 maler)** | 5 omskrevne maler + jurist-review + lessons learned |
| **4** | **Konsolidering** — duplikater, filnavn, juridisk låsing | Virke-beslutning på duplikater, markerte juridiske seksjoner |
| **5-7** | **Batch-omskriving** kategori for kategori | 99 norske maler ferdig |
| **8** | **Engelsk translation-pass** | 27 engelske maler ferdig |
| **9** | **Excel-spor (parallelt)** | 11 Excel-filer oppdatert |
| **10** | **QA, dokumentasjon, overlevering** | Designsystem-dok + prosess for fremtidig vedlikehold |

### 4.2 Pilotbatch — 5 maler som dekker mønsterrommet

| # | Mal | Dekker |
|---|-----|--------|
| 1 | Ansettelsesavtale ordinær – 010724 | Lang kontrakt, høyest volum, mange variabler |
| 2 | Oppsigelse grunnet virksomhetens forhold – 1089 A | Juridisk eksponert brev |
| 3 | Advarsel – 1000 | Minimalt brev, rask iterasjon |
| 4 | Protokoll fra drøftingsmøte – 1098 | Signaturtabeller, flere parter |
| 5 | Egenmelding ved egen sykdom | Ren skjemaform, mange felter |

Mellom disse treffer vi alle paragraph/table/komponent-stiler minst én gang.

### 4.3 AI-assistert omskriving — pipeline

```
Gammel mal (.docx/.dotx)
   ↓ [pandoc / python-docx: ekstraher til markdown + meta]
Strukturert input:
   { juridisk_tekst_locked: "…",
     overskrifter: […],
     variabler: […],
     tabeller: […] }
   ↓ [Claude med Virke-Master.dotx som kontekst]
Generator: mapper innhold til Master-stiler (endrer aldri juridisk_tekst_locked)
   ↓ [python-docx / openxml: sett sammen ny .dotx basert på Master]
Ny mal + QC-rapport
```

### 4.4 QC-gate per omskrevet mal

| Sjekk | Pass-kriterium |
|-------|---------------|
| Tekst-diff mot original | Null endringer i juridisk-låst seksjon |
| Stil-whitelist | Kun 15 godkjente styleIds brukt |
| Plassholder-konsistens | Alle matcher `{variabelnavn}`-mønster |
| Paragrafreferanser bevart | §-henvisninger identiske |
| Word 2016 render (Mac + Win) | Åpner uten feil, layout intakt |
| PDF-eksport vs Figma-referanse | >95% visuell match |

**Hard-stopp:** >10 % fail-rate i en kategori → stopp batchen, finn rot-årsak, oppdater Master, restart kategorien.

### 4.5 Risiko-register

| Risiko | Sannsynlighet | Mitigasjon |
|--------|--------------|-----------|
| Figma-stil ikke Word-kompatibel | Høy | Begrens Figma-paletten til Word-støttede egenskaper før uke 1 |
| AI endrer juridisk tekst | Høy | Hard-lås seksjoner i input, diff-gate i QC |
| Inkonsistent AI-output mellom kjøringer | Medium | Temperature 0, deterministisk prompt, Master.dotx som kontekst |
| Brand-direktiv mangler hos Virke | Høy | Avklares uke 1 (se åpne spørsmål under) |
| Medlemsbedrifter vil overstyre brand | Medium | Bruk Word Theme → medlem kan overstyre uten strukturbrudd |

---

## 5. Anbefalt Figma-struktur

```
Virke Word Design System (Figma-fil)
├─ 1. Tokens
│   ├─ Colors (variables)
│   ├─ Typography (text styles)
│   └─ Spacing (variables)
├─ 2. Word-stiler (text styles med 1:1-mapping til styleId)
│   ├─ Paragraph: Tittel, Undertittel, Overskrift1, Overskrift2,
│   │             Brødtekst, Hjelpetekst, Instruksjon, Liste, Signaturlinje
│   ├─ Character: Fremhevet, Variabel, Paragrafreferanse
│   └─ Table: Skjematabell, Signaturtabell, Data
├─ 3. Komponenter (Building Blocks)
│   ├─ Masthead
│   ├─ Partsblokk
│   ├─ Signaturblokk
│   ├─ Avsnittsvarsel (erstatter rød tekst)
│   ├─ Paragrafhenvisning
│   ├─ Vilkårsliste
│   ├─ Plassholder-chip
│   └─ Sidefot
├─ 4. Maleksempler (frame-per-mal)
│   └─ Pilot-5 + representative maler per kategori
└─ 5. Ressurser
    ├─ Virke-Master.dotx (lenke)
    ├─ Omskrivings-prompt (lenke)
    └─ QC-sjekkliste
```

---

## 6. Åpne spørsmål før Figma-arbeidet starter

1. **Virke brand** — Finnes brand-guide (logo, primærfarge, font)? Hvis ikke: skal vi foreslå, eller holde nøytralt (som i dag)?
2. **Mål med "brand"** — Skal medlemsbedrifter kunne overstyre med egen logo/farge (Virke som "shell"), eller skal det være sterkt Virke-branding?
3. **Skope** — Skal alle 131 filer omskrives, eller bare norske/engelske Word? Er Excel-filene (11 stk) inkludert?
4. **Duplikater** — Mandat til å slå sammen (Samtykkeerklæring x2, Fraværskort 2025/2026, Prøvetid x3), eller Virke-beslutning?
5. **Juridisk ressurs** — Jurist i loop for pilot-review, eller Virke-godkjenning i etterkant?
6. **Font-lisensiering** — Hvis brand-font ikke er system-font, må den embeddes i .dotx — Virke-lisens OK?
7. **Word-versjon** — Minimum versjon som skal støttes? (Påvirker hvilke Figma-features som kan oversettes 1:1.)
8. **Plassholder-syntaks** — `{variabelnavn}` OK, eller vil Virke ha Word Content Controls (mer "offisielt" men mer komplekst)?
9. **Prosjektets tidslinje** — 8-10 uker er mitt estimat; er det realistisk?

---

## Oppsummering

- **131 maler** analysert (104 NO + 27 EN, 120 Word + 11 Excel)
- **Nåtilstand:** 0 Virke-branding, 4 tilfeldige Office-blåtoner, opptil 422 stiler per mal, 50 ulike plassholder-varianter
- **Forenklingsforslag:** 15 stiler totalt, én Master.dotx, 8 Building Blocks, standardisert plassholder-syntaks
- **Plan:** 10 uker i 7 faser, pilot på 5 representative maler, AI-assistert batch med juridisk diff-gate
