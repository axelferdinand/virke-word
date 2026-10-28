from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import json,shutil,hashlib
root=Path('/Users/afg/Nettsteder/virke-word');qa=root/'Word-QA-2026-10-27';base=root/'Virke-med-komponentguide-2026-10-21';dest=root/'Virke-revidert-2026-10-27';dest.mkdir(exist_ok=True)
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';ns={'w':W};q=lambda x:'{'+W+'}'+x
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
files=sorted((qa/'kandidater').glob('[0-9]*.docx'))+[qa/'kandidater/Virke_master_test.dotx']
assert len(files)==18
records=[]
for p in files:
 with ZipFile(p) as z:
  parts={n:z.read(n) for n in z.namelist()};settings=E.fromstring(parts['word/settings.xml']);external=[]
  for n,raw in parts.items():
   if n.endswith(('.xml','.rels')):E.fromstring(raw)
   if n.endswith('.rels'):
    for r in E.fromstring(raw):
     if r.get('TargetMode')=='External':external.append({'part':n,'target':r.get('Target')})
  assert not settings.findall('w:attachedTemplate',ns)
  assert not settings.findall('w:linkStyles',ns)
  assert not external,(p,external)
  if p.suffix=='.docx':assert settings.find('w:updateFields',ns).get(q('val'))=='false'
  g=E.fromstring(parts['word/glossary/document.xml']);names=g.xpath('//w:docPart[w:docPartPr/w:types/w:type[@w:val="autoExp"]]/w:docPartPr/w:name/@w:val',namespaces=ns)
  assert len(names)==21,(p,len(names))
  styles=E.fromstring(parts['word/styles.xml'])
  for sid,b,a in [('Overskrift1','450','270'),('Overskrift3','360','180')]:
   sp=styles.find('w:style[@w:styleId="'+sid+'"]/w:pPr/w:spacing',ns);assert sp.get(q('before'))==b and sp.get(q('after'))==a
  with ZipFile(base/p.name) as old:
   opaque=[n for n in parts if any(x in n for x in ['/charts/','/media/','/embeddings/','/theme/'])]
   assert all(parts[n]==old.read(n) for n in opaque)
  text=' '.join(E.fromstring(parts['word/document.xml']).xpath('//w:t/text()',namespaces=ns))
  assert not any(v in text for v in ['QA-TEST','QA TEST','27okt','Etter signatur test','Anne Kristine'])
  records.append({'file':p.name,'sha256':sha(p),'attached_template':False,'external_relationships':external,'auto_update_styles':False,'auto_update_fields':False,'autotext_components':len(names),'style_count':len(styles),'opaque_parts_unchanged':len(opaque)})
 shutil.copy2(p,dest/p.name)
shutil.copytree(base/'Fonter',dest/'Fonter',dirs_exist_ok=True)
cpath=root/'docs/brukerveiledning-master.json';content=json.loads(cpath.read_text());content['master']=str((dest/'Virke_master_test.dotx').relative_to(root));cpath.write_text(json.dumps(content,ensure_ascii=False,indent=2)+'\n')
(qa/'final-package-audit.json').write_text(json.dumps(records,ensure_ascii=False,indent=2)+'\n')
(dest/'START_HER.txt').write_text('''VIRKE – MASTER OG KOMPONENTGUIDE
Revidert 27. oktober 2026. Veiledning 2.1.

1. Installer de to Clan-fontene fra Fonter og start Word på nytt.
2. Les Virke_mastermal_brukerveiledning.pdf. DOCX-utgaven er redigerbar
   og viser alle 21 komponentene med praktiske forklaringer.
3. Lag nye dokumenter fra Virke_master_test.dotx når du vil bruke
   masterens stiler og sette inn Virke-blokker. Lagre resultatet som DOCX.
4. Bruk kopier av de nummererte eksempeldokumentene når et ferdig oppsett
   passer. De 17 DOCX-filene viser til sammen 18 eksempelsider.

Eksempeldokumentene er selvstendige: formateringen er lagret i hver fil,
maltilknytningen er fjernet, og automatisk stil- og feltoppdatering ved
åpning er slått av. De kan flyttes til Virkes Microsoft 365-arkiv uten
å finne igjen masterfilens plassering. Oppdater innholdsfortegnelsen
manuelt etter innholdsendringer. Det er ikke nødvendig å tilknytte
masteren på nytt for å bruke de ferdige eksemplene.

Masteren brukes til vedlikehold og opprettelse av nye oppsett.
Arbeidsmåtene er kontrollert i Word for Mac. Bruk Word på skrivebordet.
Fonter må være installert; de lastes ikke automatisk fra mappen.

Notat_til_Emil.txt og TESTRESULTATER.txt dokumenterer denne revisjonen.
Pakken er master- og eksempelgrunnlaget; oppdateringen av de 40
Styres-bruksmalene er en egen leveransefase.
''')
(dest/'Notat_til_Emil.txt').write_text('''Til Emil – revisjon 27. oktober 2026

Tilbakemeldingene i Figma og Feedback.pdf er gjennomgått. H1/H3-avstand,
fete tabellhoder ved vanlig innsetting, cellebalanse og avstand etter
ferdige tabell- og signaturblokker er oppdatert og prøvd i Word.
Virke-blokker er gjort tilgjengelig under Hjem og på Virke-fanen.
De 17 eksempeldokumentene er oppdatert, kontrollert og frikoblet fra
masterfilen. Komponentguiden 2.1 viser alle 21 redigerbare byggeblokker.

Tre praktiske avgrensninger er dokumentert etter forsøk i Word for Mac:
• Rosa punkttegn fås stabilt med stilen Punktliste. Bindestrek + mellomrom
  og Words vanlige punktlisteknapp bruker fortsatt Words egne standarder.
• Virke-fortegnelsen vises ikke som et eget valg i Mac-versjonens vanlige
  galleri. Den finnes via Virke-blokker, og en synlig Virke-knapp åpner
  dialogen for innholdsfortegnelse med Fra mal og tre nivåer.
• Den faste luften etter tabeller følger Virke-blokkene. En vilkårlig ny
  tabell fra Sett inn > Tabell får korrekte tabellhoder og cellemarger,
  men Word har ikke en tilsvarende tabellstil-egenskap for luft utenfor
  en tabell som står i vanlig tekstflyt. Det er ikke lovet automatisk
  18 pt etter alle vilkårlige nyinnsatte tabeller.

De tidligere kosmetiske TOC-avvikene består: øverste skillelinje og
svarte sidetall på første nivå etter full oppdatering. Innhold, sidetall
og klikkbare henvisninger fungerer. Tidligere forsøk på å endre dette
påvirket andre deler av fortegnelsen; disse variantene er ikke brukt.

Dette er kontrollert i Word for Mac. Windows og Word på nett er ikke
verifisert i denne revisjonen. Designvurderingene er holdt utenfor den
praktiske brukerveiledningen.
''')
print('Staged',len(records),'documents; master',sha(dest/'Virke_master_test.dotx'))
