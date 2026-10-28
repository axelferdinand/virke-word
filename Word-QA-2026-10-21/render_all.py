from pathlib import Path
from concurrent.futures import ThreadPoolExecutor,as_completed
import subprocess,json
ROOT=Path('/Users/afg/Nettsteder/virke-word')
PY='/Users/afg/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3'
RENDER='/Users/afg/.codex/plugins/cache/openai-primary-runtime/documents/26.1015.12040/skills/documents/render_docx.py'
def run(p):
 out=ROOT/'Word-QA-2026-10-21/render'/p.stem
 r=subprocess.run([PY,RENDER,str(p),'--output_dir',str(out),'--emit_pdf','--verbose'],capture_output=True,text=True)
 out.mkdir(parents=True,exist_ok=True);(out/'render.log').write_text(r.stdout+r.stderr)
 return {'file':p.name,'exit_code':r.returncode,'pages':len(list(out.glob('page-*.png'))),'directory':str(out.relative_to(ROOT))}
files=sorted(p for p in (ROOT/'Virke-revidert-2026-10-21').iterdir() if p.suffix in ('.docx','.dotx'))
results=[]
with ThreadPoolExecutor(max_workers=2) as pool:
 for f in as_completed([pool.submit(run,p) for p in files]):
  result=f.result();results.append(result);print(json.dumps(result),flush=True)
(ROOT/'Word-QA-2026-10-21/render-manifest.json').write_text(json.dumps(results,indent=2)+'\n')
