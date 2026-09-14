import base64
import io
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

root = Path(__file__).parent
payload = base64.b64decode((root / 'export-v8.b64').read_text())
ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
# Read-only visual snapshot: keep Google's cached values to avoid recalculating
# Google-specific formulas in the local renderer. Never upload this copy.
with zipfile.ZipFile(io.BytesIO(payload)) as source:
    with zipfile.ZipFile(root / 'preview-cache.xlsx', 'w', zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            data = source.read(name)
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                xml = ET.fromstring(data)
                for cell in xml.iter(ns + 'c'):
                    for formula in list(cell.findall(ns + 'f')):
                        cell.remove(formula)
                # Expand rows only in the local QA copy to inspect product names.
                for row in xml.iter(ns + 'row'):
                    row.attrib.pop('hidden', None)
                data = ET.tostring(xml, encoding='utf-8', xml_declaration=True)
            target.writestr(name, data)
print('Read-only cached preview prepared')
