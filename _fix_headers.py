from pathlib import Path
path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\templates\relatorio_programacao.html")
text = path.read_text(encoding='utf-8')
text = text.replace(
    """                            <th>Saldo</th>
                            <th>AÃ§Ã£o</th>
""",
    """                            <th>Saldo</th>
                            <th>PrevisÃ£o</th>
                            <th>AÃ§Ã£o</th>
"""
)
path.write_text(text, encoding='utf-8')
print('ok')
