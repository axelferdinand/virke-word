from pathlib import Path
from PIL import Image,ImageDraw
p=Path('Word-QA-2026-10-27/eksempel-render')
files=sorted(p.glob('*/page-*.png'))
out=Path('Word-QA-2026-10-27/eksempel-kontaktark');out.mkdir(exist_ok=True)
for start in range(0,len(files),2):
 canvas=Image.new('RGB',(1800,1320),'#dddddd')
 for j,f in enumerate(files[start:start+2]):
  im=Image.open(f).convert('RGB');im.thumbnail((880,1260));canvas.paste(im,(j*900+(900-im.width)//2,40))
  ImageDraw.Draw(canvas).text((j*900+12,12),f.parent.name+' '+f.stem,fill='black')
 canvas.save(out/f'{start+1:02d}.png')
print(len(files),'pages')
