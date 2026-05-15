from pathlib import Path
import zipfile, shutil, xml.etree.ElementTree as ET

src = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\csv\O.P.xlsx")
dst = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\csv\O.P.sanitized.xlsx")
shutil.copy2(src, dst)

with zipfile.ZipFile(dst, 'r') as zin:
    files = zin.namelist()
    data = {name: zin.read(name) for name in files}

ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
tree = ET.fromstring(data['xl/workbook.xml'])
wbviews = tree.find('main:bookViews', ns)
changed = False
if wbviews is not None:
    for bv in list(wbviews):
        wbviews.remove(bv)
        changed = True
    if len(wbviews) == 0:
        tree.remove(wbviews)
        changed = True
    data['xl/workbook.xml'] = ET.tostring(tree, encoding='utf-8', xml_declaration=True)

with zipfile.ZipFile(dst, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, content in data.items():
        zout.writestr(name, content)

print('sanitized', dst, 'changed', changed)
