from pathlib import Path

path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\templates\relatorio_programacao.html")
text = path.read_text(encoding='utf-8')

# Add Previsao/Atraso headers (both tables) and remove Atraso column label later
text = text.replace(
    """                            <th>Saldo</th>
                            <th>AÃ§Ã£o</th>
""",
    """                            <th>Saldo</th>
                            <th>PrevisÃ£o</th>
                            <th>Atraso</th>
                            <th>AÃ§Ã£o</th>
"""
)

# Add previsao/atraso cells
text = text.replace(
    """                            <td class=\"num ok\">{{ item.saldo }}</td>
                            <td class=\"acao\">\n""",
    """                            <td class=\"num ok\">{{ item.saldo }}</td>
                            <td class=\"num\">{{ item.previsao or '-' }}</td>
                            <td class=\"num\">{{ 'Sim' if item.atraso else 'NÃƒÂ£o' }}</td>
                            <td class=\"acao\">\n"""
)

text = text.replace(
    """                            <td class=\"num bad\">{{ item.saldo }}</td>
                            <td class=\"acao\">\n""",
    """                            <td class=\"num bad\">{{ item.saldo }}</td>
                            <td class=\"num\">{{ item.previsao or '-' }}</td>
                            <td class=\"num\">{{ 'Sim' if item.atraso else 'NÃƒÂ£o' }}</td>
                            <td class=\"acao\">\n"""
)

# Ensure agenda attribute exists
text = text.replace(
    "data-ops='{{ item.ops | tojson | safe }}' data-secao=",
    "data-ops='{{ item.ops | tojson | safe }}' data-agenda='{{ item.agenda | tojson | safe }}' data-secao=",
)

# Update colspan
text = text.replace("colspan=\"9\"", "colspan=\"9\"")

# Replace JS block
old_js = """                const ops = JSON.parse(row.getAttribute('data-ops') || '[]');
                if (!ops.length) {
                    infoConteudo.innerHTML = '<div class=\\"info-row\\">Sem OP programada</div>';
                } else {
                    const linhas = ops.map(o => (
                        `<div class=\\"info-row\\"><b>OP:</b> ${o.op} | <b>CriaÃ§Ã£o:</b> ${o.data_emissao}</div>`
                    )).join('');
                    infoConteudo.innerHTML = linhas;
                }
                painelInfo.classList.remove('hidden');
"""

new_js = """                const ops = JSON.parse(row.getAttribute('data-ops') || '[]');
                const agenda = JSON.parse(row.getAttribute('data-agenda') || '[]');

                const agendaHtml = agenda.length
                    ? agenda.map(a => (
                        `<div class=\\"info-row\\"><b>Data:</b> ${a.data} | <b>Qtd:</b> ${a.quantidade}</div>`
                    )).join('')
                    : '<div class=\\"info-row\\">Sem previsÃ£o de faturamento</div>';

                const opsHtml = ops.length
                    ? ops.map(o => (
                        `<div class=\\"info-row\\"><b>OP:</b> ${o.op} | <b>CriaÃ§Ã£o:</b> ${o.data_emissao} | <b>Pendente:</b> ${o.pendente}</div>`
                    )).join('')
                    : '<div class=\\"info-row\\">Sem OP programada</div>';

                infoConteudo.innerHTML = `
                    <div class=\\"info-block\\">
                        <div class=\\"info-subtitle\\">PrevisÃ£o por datas</div>
                        ${agendaHtml}
                    </div>
                    <div class=\\"info-block\\">
                        <div class=\\"info-subtitle\\">OPs programadas</div>
                        ${opsHtml}
                    </div>
                `;
                painelInfo.classList.remove('hidden');
"""

text = text.replace(old_js, new_js)

path.write_text(text, encoding='utf-8')
print('ok')
