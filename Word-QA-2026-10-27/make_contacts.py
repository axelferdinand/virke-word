from PIL import Image, ImageOps, ImageDraw
from pathlib import Path
for folder in ['guide-word-render','guide-render-v4']:
 p=Path('Word-QA-2026-10-27')/folder
 files=sorted(p.glob('page-*.png'), key=lambda x:int(x.stem.split('-')[-1]))
 for start in range(0,len(files),2):
  canvas=Image.new('RGB',(1800,1320),'#dddddd')
  for j,f in enumerate(files[start:start+2]):
   im=Image.open(f).convert('RGB');im.thumbnail((880,1260));canvas.paste(im,(j*900+(900-im.width)//2,40))
   ImageDraw.Draw(canvas).text((j*900+20,12),f'{folder} — page {start+j+1}',fill='black')
  canvas.save(p/f'review-{start+1:02d}-{min(start+2,len(files)):02d}.png')
