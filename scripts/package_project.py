"""Create a complete, allowlisted archive. Never includes temporary/cache/credential files."""
import hashlib
import json
import zipfile
from project_paths import ROOT,local

def main():
    top=['NickelOps.xlsx','README.md','data_dictionary.md','test_report.md','case_study.pdf','case_study.md','demo_script.md','requirements.txt','.gitignore','.gitattributes']
    directories=['data','evaluation','exports','scripts','reports','previews']
    files=[local(x) for x in top]
    for folder in directories:
        for p in sorted(local(folder).rglob('*')):
            if p.is_symlink():raise RuntimeError('Symlink not allowed in archive')
            if p.is_file() and p.name not in ['delivery_checksums.json','debug_engine.mjs','workbook_input.json'] and '__pycache__' not in p.parts and p.suffix in ['.py','.mjs','.csv','.json','.ndjson','.png']:
                if not p.resolve().is_relative_to(ROOT):raise RuntimeError('Boundary violation')
                files.append(p)
    manifest={str(p.relative_to(ROOT)).replace('\\','/'):hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
    local('reports/delivery_checksums.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
    files.append(local('reports/delivery_checksums.json'))
    archive=local('NickelOps-Complete-Project.zip')
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
        for p in files:z.write(p,p.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(archive) as z:
        bad=z.testzip()
        if bad:raise RuntimeError('ZIP integrity failed: '+bad)
        if any(x.startswith('tmp/') or '__pycache__' in x or '..' in x.split('/') for x in z.namelist()):raise RuntimeError('Unexpected ZIP member')
        print(json.dumps(dict(archive=archive.name,files=len(z.namelist()),bytes=archive.stat().st_size,integrity='PASS')))

if __name__=='__main__':main()
