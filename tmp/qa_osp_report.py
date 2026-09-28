from pathlib import Path
import subprocess
from pypdf import PdfReader

root = Path(r'C:\Users\Administrator\Documents\New project')
pdf = root / 'outputs/osp_sourcing_report_20260928/OSP_동결건조_소싱_중간보고_20260928.pdf'
out = root / 'tmp/osp_sourcing_report_20260928_render'
out.mkdir(parents=True, exist_ok=True)
poppler = Path(r'C:\Users\CHOIIH\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin\pdftoppm.exe')
subprocess.run([str(poppler), '-png', '-r', '130', str(pdf), str(out / 'page')], check=True)
document = PdfReader(pdf)
for index, page in enumerate(document.pages):
    text = page.extract_text()
    print(f'PAGE {index+1}: {text[:250]!r}')
    print(f'END: {text[-400:]!r}')
print('Pages:', len(document.pages))
