import os, subprocess, json
from pathlib import Path
root=Path('/Users/afg/Nettsteder/virke-word')
py='/Users/afg/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3'
render='/Users/afg/.codex/plugins/cache/openai-primary-runtime/documents/26.1026.10807/skills/documents/render_docx.py'
env=dict(os.environ,FONTCONFIG_FILE=str(root/'Word-QA-2026-10-21/fonts.conf'))
results=[]
for p in sorted((root/'Word-QA-2026-10-28/kandidater').glob('[0-9]*.docx')):
 out=root/'Word-QA-2026-10-28/eksempel-render'/p.stem
 r=subprocess.run([py,render,str(p),'--output_dir',str(out),'--emit_pdf'],env=env,capture_output=True,text=True)
 results.append({'file':p.name,'returncode':r.returncode,'pages':len(list(out.glob('page-*.png')))})
 if r.returncode:print(r.stderr)
 print(p.name,results[-1]['pages'],flush=True)
(root/'Word-QA-2026-10-28/eksempel-render.json').write_text(json.dumps(results,indent=2))
