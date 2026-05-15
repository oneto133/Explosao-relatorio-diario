from pathlib import Path
path = Path(r"C:\Users\pc\Desktop\Scraping_Explosão\templates\relatorio_programacao.html")
text = path.read_text(encoding='utf-8')

insert_after = """        function limparFiltro() {
            filtroStatus.value = 'todos';
            filtroSecao.value = '';
            aplicarFiltro();
        }

"""

addon = """        const modalSequenciar = document.getElementById('modalSequenciar');
        const opAtual = document.getElementById('opAtual');
        const listaOperadores = document.getElementById('listaOperadores');
        const linhaSugestao = document.getElementById('linhaSugestao');

        let opSelecionada = null;
        let linhaSelecionada = null;

        function determinarLinha(secao, descricao) {
            const s = (secao || '').toLowerCase();
            const d = (descricao || '').toLowerCase();
            if (s.startsWith('deslizante')) return 'Deslizante';
            if (d.includes('new bv')) return 'New BV';
            if (s.includes('basculante')) return 'Basculante';
            return '';
        }

        async function abrirSequenciamento(row) {
            const ops = JSON.parse(row.getAttribute('data-ops') || '[]');
            if (!ops.length) {
                alert('Sem OP para sequenciar');
                return;
            }

            const pendente = ops.find(o => !o.sequenciado) || ops[0];
            opSelecionada = pendente.op;

            const secao = row.getAttribute('data-secao') || '';
            const desc = row.getAttribute('data-desc') || '';
            linhaSelecionada = determinarLinha(secao, desc);

            opAtual.innerText = opSelecionada;
            listaOperadores.innerHTML = '';

            if (!linhaSelecionada) {
                linhaSugestao.innerText = 'Linha não definida para este item.';
                modalSequenciar.classList.remove('hidden');
                return;
            }

            linhaSugestao.innerText = `Linha sugerida: ${linhaSelecionada}`;

            const res = await fetch(`/operadores/${linhaSelecionada}`);
            const operadores = await res.json();

            if (!operadores.length) {
                listaOperadores.innerHTML = '<p class="erro">Nenhum operador para esta linha</p>';
            } else {
                operadores.forEach(op => {
                    listaOperadores.innerHTML += `
                        <label class="radio-row">
                            <span>${op}</span>
                            <input type="radio" name="operador" value="${op}">
                        </label>
                    `;
                });
            }

            modalSequenciar.classList.remove('hidden');
        }

        function fecharSequenciamento() {
            modalSequenciar.classList.add('hidden');
        }

        async function confirmarSequenciamento() {
            if (!opSelecionada || !linhaSelecionada) {
                alert('Linha não definida');
                return;
            }
            const operador = document.querySelector('input[name="operador"]:checked');
            if (!operador) {
                alert('Selecione um operador');
                return;
            }

            await fetch('/sequenciar', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    op: opSelecionada,
                    linha: linhaSelecionada,
                    operador: operador.value
                })
            });

            fecharSequenciamento();
        }

"""

if insert_after not in text:
    raise SystemExit('marker not found')
text = text.replace(insert_after, insert_after + addon)

# Add click handler for btn-seq
old = """        document.querySelectorAll('tr.row-click').forEach(row => {
            row.addEventListener('click', (e) => {
                if (e.target.classList.contains('btn-seq')) return;
"""
new = """        document.querySelectorAll('.btn-seq').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const row = e.target.closest('tr');
                if (row) abrirSequenciamento(row);
            });
        });

        document.querySelectorAll('tr.row-click').forEach(row => {
            row.addEventListener('click', (e) => {
                if (e.target.classList.contains('btn-seq')) return;
"""
if old not in text:
    raise SystemExit('old block not found')
text = text.replace(old, new)

path.write_text(text, encoding='utf-8')
print('ok')
