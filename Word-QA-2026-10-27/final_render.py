import os,subprocess,json
from pathlib import Path
from PIL import Image,ImageDraw,ImageChops
root=Path('/Users/afg/Nettsteder/virke-word');qa=root/'Word-QA-2026-10-27'
py='/Users/afg/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3'
render='/Users/afg/.codex/plugins/cache/openai-primary-runtime/documents/26.1025.11240/skills/documents/render_docx.py'
env=dict(os.environ,FONTCONFIG_FILE=str(root/'Word-QA-2026-10-21/fonts.conf'))
out=qa/'guide-final-word-render';out.mkdir(exist_ok=True)
old=Image.open(qa/'guide-word-render/page-01.png');dpi=round(old.width*72/595.32)
subprocess.run(['/Users/afg/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdftoppm','-png','-r',str(dpi),str(qa/'testkopier/Komponentguide-2.1-final.pdf'),str(out/'page')],check=True)
diff=[]
files=sorted(out.glob('page-*.png'))
for i,f in enumerate(files,1):
 a=Image.open(f).convert('RGB');b=Image.open(qa/'guide-word-render'/f.name).convert('RGB')
 same=a.size==b.size and ImageChops.difference(a,b).getbbox() is None
 diff.append({'page':i,'pixel_identical_to_reviewed':same,'size':a.size})
for start in range(0,len(files),2):
 canvas=Image.new('RGB',(1800,1320),'#dddddd')
 for j,f in enumerate(files[start:start+2]):
  im=Image.open(f).convert('RGB');im.thumbnail((880,1260));canvas.paste(im,(j*900+(900-im.width)//2,40));ImageDraw.Draw(canvas).text((j*900+20,12),f'Final Word PDF – page {start+j+1}',fill='black')
 canvas.save(out/f'review-{start+1:02d}.png')
(qa/'guide-final-pixel-comparison.json').write_text(json.dumps(diff,indent=2));print('PDF comparison',diff,flush=True)
for n in ['09','13','16']:
 p=next((qa/'kandidater').glob(n+'*.docx'));dest=qa/'eksempel-final-render'/p.stem
 r=subprocess.run([py,render,str(p),'--output_dir',str(dest),'--emit_pdf'],env=env,capture_output=True,text=True)
 print(p.name,r.returncode,len(list(dest.glob('page-*.png'))),flush=True)
 if r.returncode:print(r.stderr)
