// ============================================================
// charts.js — Auditoria de Folha | Igarapé Digital
// Renderização com Chart.js 4.4.0 UMD
// ============================================================

if (typeof ChartDataLabels !== "undefined") {
  Chart.register(ChartDataLabels);
  Chart.defaults.set("plugins.datalabels", { display: false });
}

const CHART_REGISTRY = {};
function _destroyChart(id) { if (CHART_REGISTRY[id]) { CHART_REGISTRY[id].destroy(); delete CHART_REGISTRY[id]; } }
function _reg(id, chart) { CHART_REGISTRY[id] = chart; return chart; }
function _canvas(id) { _destroyChart(id); return document.getElementById(id); }

const FONT_DEF = { size: 11, family: "'Roboto', sans-serif" };
const FONT_SM  = { size: 10, family: "'Roboto', sans-serif" };

// Badge branco para rótulos em linhas
const BADGE = { backgroundColor: "rgba(255,255,255,0.88)", borderRadius: 3, padding: {top:1,bottom:1,left:3,right:3} };

function _baseOptions(extra = {}) {
  return Object.assign({
    animation: { duration: 180 }, responsive: true, maintainAspectRatio: false,
    plugins: { legend: { labels: { font: FONT_DEF, color: TEXTO_PRINCIPAL, boxWidth: 14, padding: 12 } } }
  }, extra);
}

function _semDados(id, msg) {
  _destroyChart(id);
  const el = document.getElementById(id); if (!el) return;
  const ctx = el.getContext("2d");
  el.width = el.offsetWidth || 400; el.height = el.offsetHeight || 300;
  ctx.clearRect(0, 0, el.width, el.height);
  ctx.font = "14px Georgia, serif"; ctx.fillStyle = IG_VERDE_ESC;
  ctx.textAlign = "center"; ctx.textBaseline = "middle";
  ctx.fillText(msg, el.width/2, el.height/2);
}

// ---- 1. Evolução mensal ----
function renderEvolucaoMensal(canvasId, resumo, mesAlvo, mesesBaseline) {
  const cv = _canvas(canvasId); if (!cv) return;
  const labels = resumo.map(r => r.Mes), idxAlvo = labels.indexOf(mesAlvo);
  const corPGTO = labels.map((_, i) => i === idxAlvo ? IG_AZUL_ESC : IG_AZUL);
  const corDESC = labels.map((_, i) => i === idxAlvo ? "#5A0030"   : IG_VINHO);

  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels, datasets: [
      { label: "PGTO", data: resumo.map(r => r.PGTO), backgroundColor: corPGTO, yAxisID: "y", order: 2 },
      { label: "DESC", data: resumo.map(r => r.DESC), backgroundColor: corDESC, yAxisID: "y", order: 3 },
      { type: "line", label: "Líquido", data: resumo.map(r => r.Liquido),
        borderColor: IG_AZUL_ESC, pointBackgroundColor: labels.map((_, i) => i === idxAlvo ? "#FF4444" : IG_AZUL_ESC),
        pointBorderColor: "white", pointBorderWidth: 1.5, pointRadius: 5.5,
        borderWidth: 2.5, tension: 0.15, yAxisID: "y1", order: 1 }
    ]},
    options: _baseOptions({
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { grid: { color: "#E8E8E8" }, ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: "PGTO / DESC (R$)", font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y1: { position: "right", grid: { drawOnChartArea: false },
          ticks: { callback: fmtBRLEixo, font: FONT_SM, color: IG_AZUL_ESC },
          title: { display: true, text: "Líquido (R$)", font: FONT_SM, color: IG_AZUL_ESC } }
      },
      plugins: {
        legend: { labels: { font: FONT_SM, color: TEXTO_PRINCIPAL, boxWidth: 12, padding: 10 } },
        tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fmtBRL(ctx.parsed.y)}` } },
        datalabels: {
          display: "auto",
          formatter: (v, ctx) => v != null ? fmtBRLEixo(Math.abs(v)) : null,
          font: (ctx) => ({ size: ctx.datasetIndex < 2 ? 7 : 9, weight: "bold" }),
          anchor: "end", align: "top",
          rotation: (ctx) => ctx.datasetIndex < 2 ? -50 : 0,
          color: (ctx) => [IG_AZUL_ESC, IG_VINHO, TEXTO_PRINCIPAL][ctx.datasetIndex] || TEXTO_PRINCIPAL,
          backgroundColor: (ctx) => ctx.datasetIndex >= 2 ? "rgba(255,255,255,0.88)" : null,
          borderRadius: (ctx) => ctx.datasetIndex >= 2 ? 3 : 0,
          padding: (ctx) => ctx.datasetIndex >= 2 ? {top:1,bottom:1,left:3,right:3} : 1,
          offset: (ctx) => ctx.datasetIndex >= 2 ? 4 : 0, clamp: false
        }
      }
    })
  }));
}

// ---- 2. Conciliação ----
function renderConciliacao(canvasId, dados, mesAlvo) {
  const cv = _canvas(canvasId); if (!cv) return;
  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels: ["PGTO", "DESC (-)", "Líquido", "Verba 9950"],
      datasets: [{ data: [dados.PGTO, -dados.DESC, dados.Liquido, dados.Verba9950],
        backgroundColor: [IG_AZUL, IG_VINHO, IG_AZUL_ESC, IG_LARANJA] }] },
    options: _baseOptions({
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => fmtBRL(ctx.parsed.y) } },
        datalabels: { display: true, anchor: "end", align: "top",
          formatter: v => fmtBRLEixo(Math.abs(v)), font: { size: 10, weight: "bold" }, color: TEXTO_PRINCIPAL }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: FONT_DEF, color: TEXTO_PRINCIPAL } },
        y: { grid: { color: "#E8E8E8" }, ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO } }
      }
    })
  }));
}

// ---- 3. Distribuição de status ----
function renderDistribuicaoStatus(canvasId, liquido) {
  const cv = _canvas(canvasId); if (!cv) return;
  const cnt = {};
  for (const r of liquido) cnt[r.STATUS] = (cnt[r.STATUS] || 0) + 1;
  const sorted = Object.entries(cnt).sort((a, b) => b[1] - a[1]);
  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels: sorted.map(([s]) => s),
      datasets: [{ data: sorted.map(([,v]) => v), backgroundColor: sorted.map(([s]) => corStatus(s)) }] },
    options: _baseOptions({
      indexAxis: "y",
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => `${fmtInt(ctx.parsed.x)} funcionários` } },
        datalabels: { display: true, anchor: "end", align: "right",
          formatter: v => fmtInt(v), font: { size: 10, weight: "bold" }, color: TEXTO_PRINCIPAL, clip: false }
      },
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { grid: { display: false }, ticks: { font: FONT_DEF, color: TEXTO_PRINCIPAL } }
      }
    })
  }));
}

// ---- 4. Headcount + salário médio ----
function renderHeadcount(canvasId, salarioV20, mesAlvo) {
  const cv = _canvas(canvasId); if (!cv) return;
  const labels = salarioV20.map(r => r.Mes), idxAlvo = labels.indexOf(mesAlvo);
  const corHC = labels.map((_, i) => i === idxAlvo ? IG_AZUL_ESC : IG_AZUL_CLR);
  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels, datasets: [
      { label: "HC (verba 0020)", data: salarioV20.map(r => r.HC),
        backgroundColor: corHC, borderColor: IG_AZUL_ESC, borderWidth: 0.5, yAxisID: "y", order: 2 },
      { type: "line", label: "Salário médio", data: salarioV20.map(r => r.Salario_Medio),
        borderColor: IG_AZUL_ESC, pointBackgroundColor: labels.map((_, i) => i === idxAlvo ? "#FF4444" : IG_AZUL_ESC),
        pointBorderColor: "white", pointBorderWidth: 1.5, pointRadius: 5.5,
        borderWidth: 2.5, tension: 0.15, yAxisID: "y1", order: 1 }
    ]},
    options: _baseOptions({
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { grid: { color: "#E8E8E8" }, ticks: { callback: v => fmtInt(v), font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: "HC (quantidade)", font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y1: { position: "right", grid: { drawOnChartArea: false },
          ticks: { callback: fmtBRLEixo, font: FONT_SM, color: IG_AZUL_ESC },
          title: { display: true, text: "Salário médio (R$)", font: FONT_SM, color: IG_AZUL_ESC } }
      },
      plugins: {
        legend: { labels: { font: FONT_SM, color: TEXTO_PRINCIPAL, boxWidth: 12, padding: 10 } },
        tooltip: { callbacks: { label: ctx => ctx.datasetIndex === 0 ? `HC: ${fmtInt(ctx.parsed.y)}` : `Salário médio: ${fmtBRL(ctx.parsed.y)}` } },
        datalabels: {
          display: "auto",
          formatter: (v, ctx) => ctx.datasetIndex === 0 ? fmtInt(v) : fmtBRLEixo(v),
          font: { size: 9, weight: "bold" }, anchor: "end", align: "top",
          color: (ctx) => ctx.datasetIndex === 0 ? IG_AZUL_ESC : TEXTO_PRINCIPAL,
          backgroundColor: "rgba(255,255,255,0.85)", borderRadius: 3,
          padding: {top:1,bottom:1,left:3,right:3}, offset: 4
        }
      }
    })
  }));
}

// ---- 5. Dispersão scatter — clique navega para drill-down ----
function renderDispersao(canvasId, liquidoAnalise, mesAlvo, onClickPoint) {
  const cv = _canvas(canvasId); if (!cv) return;
  const BASE_MIN = 100;
  const valido = liquidoAnalise.filter(r => r.MEDIA_BASELINE != null && r.MEDIA_BASELINE >= BASE_MIN);
  if (!valido.length) { _semDados(canvasId, "Dados insuficientes para gráfico de dispersão."); return; }

  const criticos = valido.filter(r =>  STATUS_CRITICOS.has(r.STATUS));
  const normais  = valido.filter(r => !STATUS_CRITICOS.has(r.STATUS));
  function toPoint(r) {
    const af = window.AF_ANON && window.AF_MAP && window.AF_MAP[r.Matrícula];
    const dMat  = af ? af.mat  : r.Matrícula;
    const dNome = af ? af.nome : r.Nome;
    return { x: r.MEDIA_BASELINE, y: Math.max(r.LIQUIDO_ALVO, BASE_MIN),
      label: `${dMat} — ${dNome}`, status: r.STATUS,
      varPct: r.VAR_PCT, varAbs: r.VAR_ABS, mat: r.Matrícula, nome: r.Nome };
  }
  const normaisPoints  = normais.map(toPoint);
  const criticosPoints = criticos.map(toPoint);
  const maxVal = Math.max(...valido.map(r => Math.max(r.MEDIA_BASELINE, Math.abs(r.LIQUIDO_ALVO), 1000)));
  const logPts = v => [BASE_MIN, BASE_MIN*5, BASE_MIN*50, maxVal*2].map(x => ({ x, y: v*x }));

  _reg(canvasId, new Chart(cv, {
    type: "scatter",
    data: { datasets: [
      { label: "Normal",  data: normaisPoints,
        backgroundColor: IG_AZUL_CLR+"BB", borderColor: IG_AZUL_CLR, borderWidth: 0.5, pointRadius: 5, order: 2 },
      { label: "Crítico", data: criticosPoints,
        backgroundColor: IG_VERMELHO+"CC", borderColor: IG_VERMELHO, borderWidth: 1, pointRadius: 8, order: 1 },
      { type: "line", label: "y=x",  data: logPts(1),   borderColor: "#444",            borderDash: [6,4], borderWidth: 1.5, pointRadius: 0, fill: false, order: 3 },
      { type: "line", label: "+30%", data: logPts(1.3), borderColor: IG_AZUL_CLR+"88", borderDash: [3,3], borderWidth: 1,   pointRadius: 0, fill: false, order: 4 },
      { type: "line", label: "-30%", data: logPts(0.7), borderColor: IG_AZUL_CLR+"88", borderDash: [3,3], borderWidth: 1,   pointRadius: 0, fill: false, order: 5 }
    ]},
    options: _baseOptions({
      onClick: (evt, elements) => {
        if (!elements.length || !onClickPoint) return;
        const el = elements[0];
        if (el.datasetIndex > 1) return;
        const arr = el.datasetIndex === 0 ? normaisPoints : criticosPoints;
        if (arr[el.index]) onClickPoint(arr[el.index]);
      },
      onHover: (evt, elements) => {
        if (evt.native) evt.native.target.style.cursor =
          (elements.length && elements[0].datasetIndex <= 1) ? "pointer" : "default";
      },
      scales: {
        x: { type: "logarithmic", grid: { color: "#E8E8E8" },
          ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: "Média líquido baseline (R$)", font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { type: "logarithmic", grid: { color: "#E8E8E8" },
          ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: `Líquido ${mesAlvo} (R$)`, font: FONT_SM, color: TEXTO_SECUNDARIO } }
      },
      plugins: {
        legend: { labels: { font: FONT_SM, color: TEXTO_PRINCIPAL, boxWidth: 12, padding: 10 } },
        tooltip: { callbacks: {
          title: ctx => ctx[0]?.raw?.label || "",
          label: ctx => ctx.raw?.status
            ? [`Status: ${ctx.raw.status}`, `Var: ${fmtPct(ctx.raw.varPct)} (${fmtBRL(ctx.raw.varAbs)})`,
               onClickPoint ? "Clique para abrir drill-down" : ""]
            : []
        }}
      }
    })
  }));
}

// ---- 6. Top 15 outliers ----
function renderTopOutliers(canvasId, liquidoAnalise, mesAlvo) {
  const cv = _canvas(canvasId); if (!cv) return;
  const top = liquidoAnalise.filter(r => STATUS_CRITICOS.has(r.STATUS))
    .sort((a, b) => Math.abs(b.VAR_ABS || 0) - Math.abs(a.VAR_ABS || 0)).slice(0, 15);
  if (!top.length) { _semDados(canvasId, `Nenhum outlier crítico detectado em ${mesAlvo}.`); return; }
  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels: top.map(r => {
        const af = window.AF_ANON && window.AF_MAP && window.AF_MAP[r.Matrícula];
        return af ? `${af.mat} — ${af.nome}` : `${r.Matrícula} — ${String(r.Nome).substring(0,28)}`;
      }),
      datasets: [{ data: top.map(r => r.VAR_ABS),
        backgroundColor: top.map(r => r.VAR_ABS > 0 ? IG_VINHO : IG_AZUL_ESC) }] },
    options: _baseOptions({
      indexAxis: "y",
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => {
          const r = top[ctx.dataIndex];
          return [`${fmtBRL(ctx.parsed.x)}`, `Status: ${r.STATUS}`, `Variação: ${fmtPct(r.VAR_PCT)}`];
        }}},
        datalabels: { display: "auto", anchor: "end",
          align: ctx => ctx.dataset.data[ctx.dataIndex] >= 0 ? "right" : "left",
          formatter: v => fmtBRLEixo(v), font: { size: 9, weight: "bold" },
          color: TEXTO_PRINCIPAL, clip: false }
      },
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: `Variação ${mesAlvo} − baseline (R$)`, font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { grid: { display: false }, ticks: { font: FONT_SM, color: TEXTO_PRINCIPAL } }
      }
    })
  }));
}

// ---- 7. Verbas zeradas ----
function renderVerbasZeradas(canvasId, dados, titulo, cor) {
  const cv = _canvas(canvasId); if (!cv) return;
  if (!dados || !dados.length) { _semDados(canvasId, "Nenhuma verba regular zerada."); return; }
  const top = dados.slice(0, 14);
  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels: top.map(r => `${r.Código} — ${String(r.Descrição).substring(0,26)}`),
      datasets: [{ data: top.map(r => r.MEDIA_BASELINE), backgroundColor: cor }] },
    options: _baseOptions({
      indexAxis: "y",
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => {
          const r = top[ctx.dataIndex];
          return [`Média baseline: ${fmtBRL(ctx.parsed.x)}`, `Meses ativos: ${r.N_MESES_ATIVOS}`];
        }}},
        datalabels: { display: "auto", anchor: "end", align: "right",
          formatter: v => fmtBRLEixo(v), font: { size: 9, weight: "bold" },
          color: TEXTO_PRINCIPAL, clip: false }
      },
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: "Média baseline (R$)", font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { grid: { display: false }, ticks: { font: FONT_SM, color: TEXTO_PRINCIPAL } }
      }
    })
  }));
}

// ---- 8. Drill funcionário ----
function renderDrillFuncionario(canvasId, serie, meses, mesAlvo, mediaBaseline, labelOverride) {
  const cv = _canvas(canvasId); if (!cv) return;
  const corLinha = STATUS_CRITICOS.has(serie.STATUS) ? IG_VERMELHO : IG_AZUL;
  const idxAlvo  = meses.indexOf(mesAlvo);
  const pRadii   = meses.map((_, i) => i === idxAlvo ? 9 : 5);
  const pColors  = meses.map((_, i) => i === idxAlvo ? "#FF4444" : corLinha);
  const datasets = [{
    label: labelOverride || "Líquido mensal",
    data: meses.map(m => serie[m] || 0),
    borderColor: corLinha, backgroundColor: corLinha,
    borderWidth: 2.5, pointRadius: pRadii, pointBackgroundColor: pColors,
    pointBorderColor: "white", pointBorderWidth: 1.5, tension: 0.15, fill: false
  }];
  if (mediaBaseline != null && !isNaN(mediaBaseline)) {
    datasets.push({ label: `Baseline: ${fmtBRL(mediaBaseline)}`,
      data: meses.map(() => mediaBaseline),
      borderColor: IG_VERDE_ESC, borderDash: [6,4], borderWidth: 1.5, pointRadius: 0, fill: false,
      datalabels: { display: false }  // sem rótulos na linha de referência
    });
  }
  _reg(canvasId, new Chart(cv, {
    type: "line",
    data: { labels: meses, datasets },
    options: _baseOptions({
      plugins: {
        legend: { labels: { font: FONT_SM, color: TEXTO_PRINCIPAL, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => `${ctx.dataset.label.split(":")[0]}: ${fmtBRL(ctx.parsed.y)}` } },
        datalabels: {
          display: "auto",
          formatter: (v, ctx) => ctx.datasetIndex === 0 ? fmtBRL(v) : null,
          font: { size: 9, weight: "bold" }, anchor: "top", align: "top",
          color: corLinha, ...BADGE, offset: 4
        }
      },
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { font: FONT_SM, color: TEXTO_SECUNDARIO, maxRotation: 45 } },
        y: { grid: { color: "#E8E8E8" }, ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: "Líquido (R$)", font: FONT_SM, color: TEXTO_SECUNDARIO } }
      }
    })
  }));
}

// ---- 9. HC + valor médio por verba ----
function renderAVHC(canvasId, dados, mesAlvo, mesesBaseline, tituloVerba) {
  const cv = _canvas(canvasId); if (!cv) return;
  const labels = dados.map(r => r.Mes), idxAlvo = labels.indexOf(mesAlvo);
  const corHC = labels.map((_, i) => i === idxAlvo ? IG_AZUL_ESC : IG_AZUL_CLR);
  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels, datasets: [
      { label: "HC (matrículas únicas)", data: dados.map(r => r.HC),
        backgroundColor: corHC, yAxisID: "y", order: 2 },
      { type: "line", label: "Valor médio por matrícula", data: dados.map(r => r.Media),
        borderColor: IG_AZUL_ESC,
        pointBackgroundColor: labels.map((_, i) => i === idxAlvo ? "#FF4444" : IG_AZUL_ESC),
        pointBorderColor: "white", pointBorderWidth: 1.5, pointRadius: 5.5,
        borderWidth: 2.5, tension: 0.15, yAxisID: "y1", order: 1 }
    ]},
    options: _baseOptions({
      plugins: {
        title: { display: !!tituloVerba, text: String(tituloVerba || "").substring(0, 80),
          color: IG_AZUL_ESC, font: { size: 12, weight: "bold", family: "Georgia, serif" },
          align: "start", padding: { bottom: 8 } },
        legend: { labels: { font: FONT_SM, color: TEXTO_PRINCIPAL, boxWidth: 12, padding: 10 } },
        tooltip: { callbacks: { label: ctx => ctx.datasetIndex === 0 ? `HC: ${fmtInt(ctx.parsed.y)}` : `Valor médio: ${fmtBRL(ctx.parsed.y)}` } },
        datalabels: {
          display: "auto",
          formatter: (v, ctx) => ctx.datasetIndex === 0 ? fmtInt(v) : fmtBRLEixo(v),
          font: { size: 9, weight: "bold" }, anchor: "end", align: "top",
          color: (ctx) => ctx.datasetIndex === 0 ? IG_AZUL_ESC : TEXTO_PRINCIPAL,
          backgroundColor: "rgba(255,255,255,0.85)", borderRadius: 3,
          padding: {top:1,bottom:1,left:3,right:3}, offset: 4
        }
      },
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { grid: { color: "#E8E8E8" }, ticks: { callback: v => fmtInt(v), font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: "HC", font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y1: { position: "right", grid: { drawOnChartArea: false },
          ticks: { callback: fmtBRLEixo, font: FONT_SM, color: IG_AZUL_ESC },
          title: { display: true, text: "Valor médio (R$)", font: FONT_SM, color: IG_AZUL_ESC } }
      }
    })
  }));
}

// ---- 10. Top verbas por impacto ----
function renderTopVerbas(canvasId, topVerbas, mesAlvo, verbasExcluidas) {
  const cv = _canvas(canvasId); if (!cv) return;
  if (!topVerbas || !topVerbas.length) { _semDados(canvasId, "Nenhuma verba com dados no período."); return; }
  const cor = topVerbas.map(r => {
    if (verbasExcluidas && verbasExcluidas.has(String(r.Código).trim())) return "#B0B0B0";
    return r["Clas."] === "PGTO" ? IG_AZUL : r["Clas."] === "DESC" ? IG_VINHO : IG_AZUL_CLR;
  });
  _reg(canvasId, new Chart(cv, {
    type: "bar",
    data: { labels: topVerbas.map(r => `${r.Código} — ${String(r.Descrição).substring(0,26)} [${r["Clas."]}]`),
      datasets: [{ data: topVerbas.map(r => r.VAR_ABS), backgroundColor: cor }] },
    options: _baseOptions({
      indexAxis: "y",
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => {
          const r = topVerbas[ctx.dataIndex];
          const excl = verbasExcluidas && verbasExcluidas.has(String(r.Código).trim());
          return [`${fmtBRL(ctx.parsed.x)} (${fmtPct(r.VAR_PCT)})`, ...(excl ? ["[EXCLUÍDA DO CÁLCULO]"] : [])];
        }}},
        datalabels: { display: "auto", anchor: "end",
          align: ctx => ctx.dataset.data[ctx.dataIndex] >= 0 ? "right" : "left",
          formatter: v => fmtBRLEixo(v), font: { size: 9, weight: "bold" },
          color: TEXTO_PRINCIPAL, clip: false }
      },
      scales: {
        x: { grid: { color: "#E8E8E8" }, ticks: { callback: fmtBRLEixo, font: FONT_SM, color: TEXTO_SECUNDARIO },
          title: { display: true, text: `Variação ${mesAlvo} vs média baseline (R$)`, font: FONT_SM, color: TEXTO_SECUNDARIO } },
        y: { grid: { display: false }, ticks: { font: { size: 9, family: "'Roboto', sans-serif" }, color: TEXTO_PRINCIPAL } }
      }
    })
  }));
}
