from pathlib import Path
import zipfile, xml.etree.ElementTree as ET

path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\csv\O.P.sanitized.xlsx")

with zipfile.ZipFile(path, 'r') as zin:
    files = zin.namelist()
    data = {name: zin.read(name) for name in files}

cp_ns = {'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
         'dc': 'http://purl.org/dc/elements/1.1/',
         'dcterms': 'http://purl.org/dc/terms/',
         'dcmitype': 'http://purl.org/dc/dcmitype/',
         'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}

core_xml = data.get('docProps/core.xml')
if core_xml:
    root = ET.fromstring(core_xml)
    removed = False
    for el in list(root):
        if el.tag.endswith('lastPrinted'):
            root.remove(el)
            removed = True
    if removed:
        data['docProps/core.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, content in data.items():
        zout.writestr(name, content)

print('core sanitized')
