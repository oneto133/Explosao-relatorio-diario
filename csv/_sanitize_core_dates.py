from pathlib import Path
import zipfile, xml.etree.ElementTree as ET

path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\csv\O.P.sanitized.xlsx")

with zipfile.ZipFile(path, 'r') as zin:
    data = {name: zin.read(name) for name in zin.namelist()}

core_xml = data.get('docProps/core.xml')
if core_xml:
    root = ET.fromstring(core_xml)
    removed = False
    for el in list(root):
        tag = el.tag
        if tag.endswith('created') or tag.endswith('modified'):
            root.remove(el)
            removed = True
    if removed:
        data['docProps/core.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, content in data.items():
        zout.writestr(name, content)

print('core dates removed')
