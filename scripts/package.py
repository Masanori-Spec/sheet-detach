from pathlib import Path
import hashlib,json,zipfile,shutil
root=Path(__file__).resolve().parents[1]
out=root.parent/'sheet-detach-output';out.mkdir(exist_ok=True)
exclude={'node_modules','test-results','.git','__pycache__'}
files=sorted(p for p in root.rglob('*') if p.is_file() and not any(part in exclude for part in p.relative_to(root).parts))
manifest={'name':'SheetDetach','version':'0.1.0','status':'See docs/VERIFICATION.md for exact hosted test evidence and scope; Microsoft Excel unverified','files':[]}
archive=out/'sheet-detach-source.zip'
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as z:
 for p in files:
  rel=p.relative_to(root).as_posix();data=p.read_bytes();zi=zipfile.ZipInfo('sheet-detach/'+rel,(2026,10,4,0,0,0));zi.compress_type=zipfile.ZIP_DEFLATED;zi.external_attr=0o644<<16;z.writestr(zi,data);manifest['files'].append({'path':rel,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()})
manifest['archive']={'name':archive.name,'bytes':archive.stat().st_size,'sha256':hashlib.sha256(archive.read_bytes()).hexdigest()}
(out/'source-manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
with zipfile.ZipFile(archive) as z:
 assert z.testzip() is None
 for item in manifest['files']:assert hashlib.sha256(z.read('sheet-detach/'+item['path'])).hexdigest()==item['sha256']
for name,src in [('sheet-detach.html','dist/index.html'),('handoff.xlsx','generated/handoff.xlsx'),('ledger.csv','generated/ledger.csv'),('recipe.json','generated/recipe.json'),('report.txt','generated/report.txt')]:shutil.copyfile(root/src,out/name)
print(json.dumps({'archive':manifest['archive'],'files':len(files)},indent=2))
