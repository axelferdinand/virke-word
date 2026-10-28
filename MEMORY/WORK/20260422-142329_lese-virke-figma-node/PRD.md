---
task: Lese Virke Figma designsystem node 2639-53314
slug: 20260422-142329_lese-virke-figma-node
effort: standard
phase: complete
progress: 10/10
mode: interactive
started: 2026-04-22T12:23:29Z
updated: 2026-04-22T12:23:29Z
---

## Context

Bruker delte Figma-lenke til Virke-Design-system (nodeId 2639-53314, dev-mode) og spør "Hva ser du?". Første faktiske kontakt med Virkes eksisterende designsystem — svarer også på åpent spørsmål #1 fra forrige orienterings-PRD ("Finnes brand-guide?").

### Risks

- Figma Desktop kan ha annen fil åpen enn URL-en peker til
- Node 2639-53314 kan være dypt nestet og kreve navigering
- `get_variable_defs` kan returnere hele filens variabler (ikke nodens)

## Criteria

- [x] ISC-1: figma-use skill invokert (obligatorisk prerequisite)
- [x] ISC-2: Metadata hentet for node 2639-53314
- [x] ISC-3: Screenshot hentet for node 2639-53314
- [x] ISC-4: Variabler (farger/type/spacing) hentet fra filen
- [x] ISC-5: Node-type identifisert (frame, component, side)
- [x] ISC-6: Visuelt innhold beskrevet fra screenshot
- [x] ISC-7: Tokens-funn rapportert (antall/kategorier)
- [x] ISC-8: Koblet til Virke Word-redesign-kontekst
- [x] ISC-9: Neste naturlige steg foreslått
- [x] ISC-10: Svar levert på norsk
