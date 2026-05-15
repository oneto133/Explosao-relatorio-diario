from pathlib import Path

path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\templates\relatorio_programacao.html")
text = path.read_text(encoding='utf-8')

# Remove Atraso header and column
text = text.replace("<th>Atraso</th>\n                            ", "")

text = text.replace("<td class=\"num\">{{ 'Sim' if item.atraso else 'NÃƒÂ£o' }}</td>\n                            ", "")

# Add row-atraso class based on item.atraso
text = text.replace("<tr class=\"row-click\" data-ops=", "<tr class=\"row-click {% if item.atraso %}row-atraso{% endif %}\" data-ops=")

# Update colspan (was 9 with Atraso, now 8)
text = text.replace("colspan=\"9\"", "colspan=\"8\"")

path.write_text(text, encoding='utf-8')
print('ok')
