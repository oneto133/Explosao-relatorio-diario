from pathlib import Path
import zipfile, xml.etree.ElementTree as ET

path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\csv\O.P.sanitized.xlsx")

with zipfile.ZipFile(path, 'r') as zin:
    data = {name: zin.read(name) for name in zin.namelist()}

sheet_files = [n for n in data if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

for sf in sheet_files:
    root = ET.fromstring(data[sf])
    removed = False
    for el in list(root):
        if el.tag.endswith('pageSetup') or el.tag.endswith('pageMargins') or el.tag.endswith('printOptions'):
            root.remove(el)
            removed = True
    if removed:
        data[sf] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, content in data.items():
        zout.writestr(name, content)

print('worksheets sanitized')
