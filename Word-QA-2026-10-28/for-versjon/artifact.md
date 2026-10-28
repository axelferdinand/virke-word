# Virke komponentguide 2.1 – gjeldende dokumentkontrakt

Referanse: /Users/afg/Nettsteder/virke-word/Virke-revidert-2026-10-27/Virke_master_test.dotx
SHA-256: d2f20b2236e365aef203adfc1757d1955d4f1ffcbf43ac8628d940be9c5bf784
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
