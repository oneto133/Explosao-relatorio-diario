(function () {
  const codigoInput = document.getElementById("codigoInput");
  const btnBuscar = document.getElementById("btnBuscar");
  const maPeriodo = document.getElementById("maPeriodo");
  const forecastMeses = document.getElementById("forecastMeses");
  const showMediaMovel = document.getElementById("showMediaMovel");
  const showTendencia = document.getElementById("showTendencia");
  const showPrevLinear = document.getElementById("showPrevLinear");
  const showPrevPreditiva = document.getElementById("showPrevPreditiva");
  const statusEl = document.getElementById("status");
  const predTableBody = document.getElementById("predTableBody");
  const metaCodigo = document.getElementById("metaCodigo");
  const metaDescricao = document.getElementById("metaDescricao");
  const statMediaMensal = document.getElementById("statMediaMensal");
  const statTendencia = document.getElementById("statTendencia");
  const statVolatilidade = document.getElementById("statVolatilidade");
  const statProximoMes = document.getElementById("statProximoMes");

  let chart = null;
  let currentItem = null;

  function setStatus(msg, type) {
    statusEl.textContent = msg || "";
    statusEl.className = "status" + (type ? ` ${type}` : "");
  }

  function toMonthDate(label) {
    const [mm, yyyy] = String(label).split("/").map(Number);
    return new Date(yyyy, (mm || 1) - 1, 1);
  }

  function formatMonthLabel(date) {
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const yyyy = String(date.getFullYear());
    return `${mm}/${yyyy}`;
  }

  function addMonths(label, n) {
    const d = toMonthDate(label);
    d.setMonth(d.getMonth() + n);
    return formatMonthLabel(d);
  }

  function mean(values) {
    if (!values.length) return 0;
    return values.reduce((acc, v) => acc + v, 0) / values.length;
  }

  function stdDev(values) {
    if (!values.length) return 0;
    const m = mean(values);
    const variance = mean(values.map((v) => (v - m) ** 2));
    return Math.sqrt(variance);
  }

  function movingAverage(values, period) {
    const p = Math.max(1, period);
    const out = [];
    for (let i = 0; i < values.length; i += 1) {
      if (i + 1 < p) {
        out.push(null);
        continue;
      }
      const slice = values.slice(i + 1 - p, i + 1);
      out.push(mean(slice));
    }
    return out;
  }

  function linearRegression(values) {
    const n = values.length;
    if (n === 0) return { slope: 0, intercept: 0, history: [] };
    if (n === 1) return { slope: 0, intercept: values[0], history: [values[0]] };

    let sumX = 0;
    let sumY = 0;
    let sumXY = 0;
    let sumXX = 0;
    for (let i = 0; i < n; i += 1) {
      const x = i;
      const y = values[i];
      sumX += x;
      sumY += y;
      sumXY += x * y;
      sumXX += x * x;
    }

    const denom = (n * sumXX) - (sumX * sumX);
    const slope = denom === 0 ? 0 : ((n * sumXY) - (sumX * sumY)) / denom;
    const intercept = (sumY - (slope * sumX)) / n;
    const history = values.map((_, i) => intercept + (slope * i));
    return { slope, intercept, history };
  }

  function forecastLinear(values, horizon) {
    const n = values.length;
    const model = linearRegression(values);
    const future = [];
    for (let i = 1; i <= horizon; i += 1) {
      future.push(model.intercept + (model.slope * (n - 1 + i)));
    }
    return { model, future };
  }

  function forecastMovingAverage(values, period, horizon) {
    const p = Math.max(1, period);
    const rolling = [...values];
    const future = [];
    for (let i = 0; i < horizon; i += 1) {
      const base = rolling.slice(Math.max(0, rolling.length - p));
      const next = base.length ? mean(base) : 0;
      future.push(next);
      rolling.push(next);
    }
    return future;
  }

  function formatNumber(v) {
    if (v === null || v === undefined || Number.isNaN(v)) return "-";
    return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 }).format(v);
  }

  function renderStats(values, slope, nextPredictive) {
    const avg = mean(values);
    const vol = stdDev(values);
    statMediaMensal.textContent = formatNumber(avg);
    statVolatilidade.textContent = formatNumber(vol);
    statTendencia.textContent = slope.toFixed(2);
    statProximoMes.textContent = formatNumber(nextPredictive);
  }

  function renderPredictionTable(futureLabels, linearFuture, maFuture, predictiveFuture) {
    predTableBody.innerHTML = "";
    if (!futureLabels.length) {
      predTableBody.innerHTML = "<tr><td colspan='4'>Sem dados para previsao.</td></tr>";
      return;
    }

    for (let i = 0; i < futureLabels.length; i += 1) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${futureLabels[i]}</td>
        <td>${formatNumber(linearFuture[i])}</td>
        <td>${formatNumber(maFuture[i])}</td>
        <td>${formatNumber(predictiveFuture[i])}</td>
      `;
      predTableBody.appendChild(tr);
    }
  }

  function renderChart(item) {
    const period = Math.max(2, Number(maPeriodo.value) || 3);
    const horizon = Math.min(12, Math.max(1, Number(forecastMeses.value) || 6));
    const labels = item.serie.map((p) => p.mes);
    const values = item.serie.map((p) => Number(p.valor) || 0);

    const maHistory = movingAverage(values, period);
    const linear = forecastLinear(values, horizon);
    const maFuture = forecastMovingAverage(values, period, horizon);
    const predictiveFuture = linear.future.map((v, i) => (v * 0.65) + (maFuture[i] * 0.35));

    const futureLabels = [];
    const lastLabel = labels[labels.length - 1];
    for (let i = 1; i <= horizon; i += 1) {
      futureLabels.push(addMonths(lastLabel, i));
    }
    const allLabels = [...labels, ...futureLabels];

    const historyData = [...values, ...Array(horizon).fill(null)];
    const movingData = [...maHistory, ...Array(horizon).fill(null)];
    const trendData = [...linear.model.history, ...Array(horizon).fill(null)];
    const linearPredData = [...Array(values.length - 1).fill(null), values[values.length - 1], ...linear.future];
    const predictiveData = [...Array(values.length - 1).fill(null), values[values.length - 1], ...predictiveFuture];

    const datasets = [
      {
        label: "Historico",
        data: historyData,
        borderColor: "#204f74",
        backgroundColor: "rgba(32,79,116,0.14)",
        borderWidth: 2.4,
        pointRadius: 3,
        tension: 0.25
      },
      {
        label: `Media movel (${period})`,
        data: movingData,
        borderColor: "#cc6e2f",
        borderDash: [7, 4],
        borderWidth: 2,
        pointRadius: 0,
        tension: 0.2,
        hidden: !showMediaMovel.checked
      },
      {
        label: "Tendencia linear",
        data: trendData,
        borderColor: "#6c7d8e",
        borderDash: [2, 5],
        borderWidth: 1.8,
        pointRadius: 0,
        tension: 0,
        hidden: !showTendencia.checked
      },
      {
        label: "Previsao linear",
        data: linearPredData,
        borderColor: "#8e2a55",
        borderWidth: 2.2,
        pointRadius: 2,
        tension: 0.2,
        hidden: !showPrevLinear.checked
      },
      {
        label: "Previsao preditiva",
        data: predictiveData,
        borderColor: "#2d8b4c",
        borderWidth: 2.2,
        pointRadius: 2,
        tension: 0.2,
        hidden: !showPrevPreditiva.checked
      }
    ];

    const ctx = document.getElementById("analiseChart");
    if (chart) {
      chart.destroy();
    }

    chart = new Chart(ctx, {
      type: "line",
      data: {
        labels: allLabels,
        datasets
      },
      options: {
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: {
            position: "bottom",
            labels: { boxWidth: 18, boxHeight: 10, font: { weight: "bold" } }
          },
          tooltip: {
            callbacks: {
              label: (context) => `${context.dataset.label}: ${formatNumber(context.parsed.y)}`
            }
          }
        },
        scales: {
          x: {
            ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 12 }
          },
          y: {
            beginAtZero: true,
            ticks: {
              callback: (value) => new Intl.NumberFormat("pt-BR").format(value)
            }
          }
        }
      }
    });

    renderStats(values, linear.model.slope, predictiveFuture[0] || 0);
    renderPredictionTable(futureLabels, linear.future, maFuture, predictiveFuture);
  }

  async function buscar() {
    const codigo = codigoInput.value.trim();
    if (!codigo) {
      setStatus("Informe um codigo para pesquisar.", "error");
      return;
    }

    setStatus("Buscando dados...", "");
    try {
      const res = await fetch(`/api/analise-mes-a-mes/item?codigo=${encodeURIComponent(codigo)}`);
      const data = await res.json();
      if (!res.ok || !data.sucesso) {
        throw new Error(data.mensagem || "Falha na consulta.");
      }

      currentItem = data.item;
      metaCodigo.textContent = data.item.codigo || "-";
      metaDescricao.textContent = data.item.descricao || "-";
      renderChart(currentItem);
      setStatus("Analise carregada com sucesso.", "ok");
    } catch (err) {
      setStatus(err.message || "Erro ao buscar dados.", "error");
    }
  }

  function rerenderIfReady() {
    if (currentItem) {
      renderChart(currentItem);
    }
  }

  btnBuscar.addEventListener("click", buscar);
  codigoInput.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") buscar();
  });
  maPeriodo.addEventListener("change", rerenderIfReady);
  forecastMeses.addEventListener("change", rerenderIfReady);
  showMediaMovel.addEventListener("change", rerenderIfReady);
  showTendencia.addEventListener("change", rerenderIfReady);
  showPrevLinear.addEventListener("change", rerenderIfReady);
  showPrevPreditiva.addEventListener("change", rerenderIfReady);
})();
