from pathlib import Path
path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\templates\relatorio_programacao.html")
text = path.read_text(encoding='utf-8')
start = "const ops = JSON.parse(row.getAttribute('data-ops') || '[]');"
end = "painelInfo.classList.remove('hidden');"
if start not in text or end not in text:
    raise SystemExit('markers not found')
pre, rest = text.split(start, 1)
mid, post = rest.split(end, 1)
new_block = """const ops = JSON.parse(row.getAttribute('data-ops') || '[]');
                const agenda = JSON.parse(row.getAttribute('data-agenda') || '[]');
                const emAtraso = row.classList.contains('row-atraso');

                const agendaHtml = agenda.length
                    ? agenda.map(a => (
                        `<div class=\"info-row\"><b>Data:</b> ${a.data} | <b>Qtd:</b> ${a.quantidade}</div>`
                    )).join('')
                    : '<div class=\"info-row\">Sem previsÃ£o de faturamento</div>';

                const opsHtml = ops.length
                    ? ops.map(o => (
                        `<div class=\"info-row\"><b>OP:</b> ${o.op} | <b>CriaÃ§Ã£o:</b> ${o.data_emissao} | <b>Pendente:</b> ${o.pendente}</div>`
                    )).join('')
                    : '<div class=\"info-row\">Sem OP programada</div>';

                infoConteudo.innerHTML = `
                    <div class=\"info-block\">
                        <div class=\"info-subtitle\">PrevisÃ£o por datas</div>
                        ${agendaHtml}
                    </div>
                    <div class=\"info-block\">
                        <div class=\"info-subtitle\">OPs programadas</div>
                        ${opsHtml}
                    </div>
                    ${emAtraso ? '<div class=\"info-row info-atraso\">Em atraso</div>' : ''}
                `;
                """
text = pre + new_block + post
path.write_text(text, encoding='utf-8')
print('ok')
