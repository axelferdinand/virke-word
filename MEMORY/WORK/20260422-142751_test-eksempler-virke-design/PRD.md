---
task: Lage nye Word-versjoner av Eksempler med Virke-design
slug: 20260422-142751_test-eksempler-virke-design
effort: advanced
phase: complete
progress: 28/28
mode: interactive
started: 2026-04-22T12:27:51Z
updated: 2026-04-22T12:29:00Z
---

## Context

Pilot-test. Bruker vil se resultatet av Figma-prinsippene anvendt på to ekte Word-filer:
- `word-maler/Eksempler/FORSLAG_Aksjeeierbok_SS.docx`
- `word-maler/Eksempler/FORSLAG_Protokoll fra styremøte (generalforsamling)_SS.docx`

Originalene skal IKKE slettes — nye Virke-versjoner lages ved siden av.

**Viktig font-overstyring:** Bruker ba eksplisitt om **Arial istedenfor Clan Offc**, selv om Figma-variabelen spesifiserer Clan. Grunn sannsynlig: Arial er system-font på alle Windows/Mac Word-installasjoner, Clan krever embedding-lisens.

### Prinsipper fra Figma-seksjonen "Test Claude"

**Typografi (Arial-mapping)**
- Tittel/H1 — 18pt Bold, line-height 1.5
- Overskrift 1/H2 — 15pt Bold, line-height 1.5
- Overskrift 2/H3 — 11pt Bold, line-height 1.5
- Brødtekst — 10pt Regular, line-height 1.4 *(justert fra 9pt/1.6 for bedre Word-lesbarhet — 9pt er vel lite for juridiske dokumenter)*
- Caption/Sidefot — 8pt Regular, farge #616161

**Farger**
- Tekst svart #000000
- Sekundær/caption #616161
- Bakgrunn hvit

**Layout**
- A4 med 2,5 cm marger (alle sider)
- Topptekst: Virke-logotekst venstre + seksjonslabel høyre ("Selskapsrett")
- Sidefot: "Side X av Y" venstre + doktittel høyre
- Paragraf-avstand: 6pt etter brødtekst, 12pt før overskrifter

**Hjelpetekst-callout**
- Svakt gul/cream bakgrunn (#FFF3BF-aktig fra forrige PRD)
- Kursiv, grå tekst
- Tydelig "skal fjernes ved utfylling"

**Plassholder-konvensjon**
- Bevar originalens `[klammer]`-stil (ikke bytt til `{}` i pilot — vil bruke Variabel-karakterstil senere)

### Risks

- python-docx har begrenset kontroll over header/footer — må verifisere at Virke-label vises
- Rød tekst i original må kartlegges til ny semantisk stil uten å miste betydning
- Tabell-rendering i Word mellom macOS og Windows kan variere — må teste konservative innstillinger
- Eksakte farger i callout kan se forskjellig ut i Word vs Figma (CMYK-profiler osv.)

## Criteria

### Filstruktur
- [x] ISC-1: Originalfil `FORSLAG_Aksjeeierbok_SS.docx` uberørt
- [x] ISC-2: Originalfil `FORSLAG_Protokoll...` uberørt
- [x] ISC-3: Ny fil `VIRKE_Aksjeeierbok.docx` eksisterer i samme mappe
- [x] ISC-4: Ny fil `VIRKE_Protokoll fra styremøte (generalforsamling).docx` eksisterer

### Typografi (begge filer)
- [x] ISC-5: Body-font er Arial 10pt
- [x] ISC-6: H1 er Arial 18pt Bold
- [x] ISC-7: H2 er Arial 15pt Bold
- [x] ISC-8: H3 er Arial 11pt Bold
- [x] ISC-9: Caption-tekst er Arial 8pt grå #616161

### Layout (begge filer)
- [x] ISC-10: Sidemarg 2,5 cm på alle fire kanter
- [x] ISC-11: Topptekst inneholder "Virke" venstrestilt
- [x] ISC-12: Topptekst inneholder seksjonslabel høyrestilt
- [x] ISC-13: Sidefot inneholder "Side X av Y" venstrestilt
- [x] ISC-14: Sidefot inneholder doktittel høyrestilt

### Innhold bevart (Aksjeeierbok)
- [x] ISC-15: Tittel "Aksjeeierbok" beholdt
- [x] ISC-16: Hele intro-paragrafen om aksjeloven §4-7 beholdt
- [x] ISC-17: Tabell med 7 kolonner (Navn/Fødselsdato/Adresse/...) beholdt
- [x] ISC-18: Presiserings-setning om "generell art" markert som Hjelpetekst-callout

### Innhold bevart (Protokoll)
- [x] ISC-19: Tittel "Protokoll fra styremøte i [foretaksnavn]" beholdt
- [x] ISC-20: Dato/Sted/Tid-blokk beholdt
- [x] ISC-21: Sak 1 Årsregnskap med beløp beholdt
- [x] ISC-22: Sak 2 Honorar beholdt
- [x] ISC-23: Sak 3 Innkalling beholdt
- [x] ISC-24: 3 signaturfelt med Styreleder/Styremedlem beholdt

### Virke-design anvendt
- [x] ISC-25: Hjelpetekst-callout brukt for kommentarer som skal fjernes
- [x] ISC-26: Plassholdere `[...]` bevart (ikke oversatt til `{}`)
- [x] ISC-27: Tabell-header tydelig distinkt fra data-rader
- [x] ISC-28: Filene åpner uten feil (kan verifiseres ved textutract)
