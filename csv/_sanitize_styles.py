from pathlib import Path
import zipfile

path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\csv\O.P.sanitized.xlsx")

minimal = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'''

with zipfile.ZipFile(path, 'r') as zin:
    data = {name: zin.read(name) for name in zin.namelist()}

data['xl/styles.xml'] = minimal

with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, content in data.items():
        zout.writestr(name, content)

print('styles sanitized')
