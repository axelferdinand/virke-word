from zipfile import ZipFile
from lxml import etree as E
from pathlib import Path
p=Path('Virke-revidert-2026-10-27/02_Balance_scorecard.docx');ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
with ZipFile(p) as z:
 s=E.fromstring(z.read('word/styles.xml'))
 for sid in ['Normal','Standardtabell','Virke1','Virke2','VirkeTabelltekst','VirkeTabelloverskrift','Bildetekst']:
  n=s.find('w:style[@w:styleId="'+sid+'"]',ns)
  if n is not None:print(E.tostring(n,pretty_print=True).decode())
 print('SETTINGS',E.tostring(E.fromstring(z.read('word/settings.xml'))).decode())
 d=E.fromstring(z.read('word/document.xml'));t=d.find('.//w:tbl',ns)
 print('FIRST TABLE',E.tostring(t,pretty_print=True).decode())
