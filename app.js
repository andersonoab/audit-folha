// ============================================================
// app.js — Auditoria de Folha | Igarapé Digital
// Controlador principal (DOM + estado)
// ============================================================

const APP = (() => {
  // ---- Estado global ----
  const S = {
    df: null, mesesDetectados: [],
    empresasDisp: [], processosDisp: [],
    empresasSel: [], processosSel: [],
    checksBaseline: {},
    mesAlvo: null, mesesBaseline: [],
    excluirDesligados: true,
    configNome: "Shewhart 3-sigma + Materialidade ISA 320 (recomendado)",
    matPct: 1.0, verbasExcluidas: new Set(),
    // resultados
    resumo: null, liquido: null, liquidoAnalise: null,
    salarioV20: null, zeradosPGTO: null, zeradosDESC: null,
    metadata: {}, dfFiltrado: null,
    // descontos para revisar + modo mês único + filtro da Visão Geral
    descontosRevisar: null, modoMesUnico: false, vgBusca: "",
    // análise de verbas
    verbasAvCompleta: null, dfAvBase: null,
    codigosColabAtual: [], colabDfAtual: null,
    topVerbasDF: null, sumLiquids: new Map(),
    // tabela
    funcSelecionado: null,
    sortCol: "VAR_ABS", sortAsc: false,
    tabsDirty: {},
    // privacidade
    anonimizar: false,
    mapaAnonimo: {},
    csvBruto: null, csvNome: null,
    // conferência (revisão humana)
    conf: {}, filtroConf: "todos",
    // ordenação por clique (asc/desc) das tabelas
    confSort: { key: null, dir: 1 }, drillSort: { key: null, dir: 1 }, colabSort: { key: null, dir: 1 }, descSort: { key: null, dir: 1 }
  };

  function _sumEmployeeId(r) { return JSON.stringify([funcKey(r), S.mesAlvo]); }
  function _sumPayroll(row) {
    return RowSelection.payroll(row[`${S.mesAlvo} - Valor`] || 0, row["Clas."], row["Código"], row["Descrição"], _isLiqRow(row));
  }
  function _sumLiquid(r) {
    return { ["Líquido calculado · " + S.mesAlvo]: S.sumLiquids.get(funcKey(r)) || 0n };
  }
  function _sumView(name, rows, options) {
    const view = RowSelection.table(name);
    const page = view.prepare(rows, Object.assign({ key: _sumEmployeeId, label: S.mesAlvo, emptyGroup: 'Líquido calculado · ' + S.mesAlvo, values: _sumLiquid }, options));
    return { view, page };
  }

  // ---- Init ----
  function init() {
    _bindEvents();
    _popularComboMetodo();
    _confLoad();
    _loadFromStorage(); // carrega dados persistidos automaticamente
  }

  function _bindEvents() {
    document.getElementById("btnUpload").addEventListener("click", () => document.getElementById("fileInput").click());
    document.getElementById("fileInput").addEventListener("change", _onFileChange);
    document.querySelectorAll(".tab-btn").forEach(b => b.addEventListener("click", () => _switchTab(b.dataset.tab)));
    document.getElementById("comboMetodo").addEventListener("change", _onMetodoChange);
    document.getElementById("sliderMat").addEventListener("input", _onMatChange);
    document.getElementById("switchDesligados").addEventListener("change", e => { S.excluirDesligados = e.target.checked; });
    document.getElementById("btnReprocessar").addEventListener("click", acaoReprocessar);
    document.getElementById("btnFiltroVerbas").addEventListener("click", _abrirFiltroVerbas);
    document.getElementById("btnAplicarTab").addEventListener("click", _renderizarTabela);
    document.getElementById("btnLimparTab").addEventListener("click", _limparFiltrosTabela);
    document.getElementById("entryBusca").addEventListener("keydown", e => { if (e.key === "Enter") _renderizarTabela(); });
    document.getElementById("comboDrillCR").addEventListener("change", _renderizarDrillVerbas);
    document.getElementById("comboDrillClasse").addEventListener("change", _renderizarDrillVerbas);
    document.getElementById("comboDrillProc").addEventListener("change", _renderizarDrillVerbas);
    document.getElementById("entryDrillVerba").addEventListener("keydown", e => { if (e.key === "Enter") _renderizarDrillVerbas(); });
    document.getElementById("btnLimparDrill").addEventListener("click", _limparFiltrosDrill);
    document.getElementById("btnExportar").addEventListener("click", acaoExportarExcel);
    document.getElementById("entryBuscaVerba").addEventListener("input", _filtrarListboxVerbas);
    const cConf = document.getElementById("comboConfTab");
    if (cConf) cConf.addEventListener("change", _renderizarTabela);
    const bCE = document.getElementById("btnConfExport"); if (bCE) bCE.addEventListener("click", confExport);
    const bCI = document.getElementById("btnConfImport"); if (bCI) bCI.addEventListener("click", () => document.getElementById("confFileInput").click());
    const cFI = document.getElementById("confFileInput"); if (cFI) cFI.addEventListener("change", e => { if (e.target.files[0]) confImport(e.target.files[0]); e.target.value = ""; });
    const bCR = document.getElementById("btnConfReset"); if (bCR) bCR.addEventListener("click", confReset);
    const cf2 = document.getElementById("comboConfFiltro2"); if (cf2) cf2.addEventListener("change", _renderConferencia);
    const eb2 = document.getElementById("entryConfBusca2"); if (eb2) eb2.addEventListener("input", _renderConferencia);
    const bE2 = document.getElementById("btnConfExport2"); if (bE2) bE2.addEventListener("click", confExport);
    const bI2 = document.getElementById("btnConfImport2"); if (bI2) bI2.addEventListener("click", () => document.getElementById("confFileInput").click());
    const bR2 = document.getElementById("btnConfReset2"); if (bR2) bR2.addEventListener("click", confReset);
    const bMV = document.getElementById("btnConfMarcarVisiveis"); if (bMV) bMV.addEventListener("click", _confMarcarVisiveis);
    const hSel = document.getElementById("holFunc");
    if (hSel) hSel.addEventListener("change", e => { const [emp, m, n] = JSON.parse(e.target.value); HOL.empresa = emp; HOL.mat = m; HOL.nome = n; HOL.q = ""; HOL.onlyChk = false; _renderHolerite(); });
    const hBC = document.getElementById("holBuscaColab");
    if (hBC) hBC.addEventListener("input", e => { HOL.buscaColab = e.target.value; HOL.q = ""; HOL.onlyChk = false; _renderHolerite(); });
    const hPrev = document.getElementById("holPrev"); if (hPrev) hPrev.addEventListener("click", () => _holNav(-1));
    const hNext = document.getElementById("holNext"); if (hNext) hNext.addEventListener("click", () => _holNav(1));
    const hPrint = document.getElementById("holPrint"); if (hPrint) hPrint.addEventListener("click", () => window.print());
    _bindConfBox();
    document.getElementById("listboxVerbas").addEventListener("change", _renderDrillVerba);
    document.getElementById("btnAplicarColab").addEventListener("click", _renderizarTabelaColab);
    document.getElementById("btnLimparColab").addEventListener("click", () => {
      document.getElementById("comboColabStatus").value = "(todos)";
      document.getElementById("entryColabBusca").value = "";
      _renderizarTabelaColab();
    });
    document.getElementById("comboColabStatus").addEventListener("change", _renderizarTabelaColab);
    document.getElementById("entryColabBusca").addEventListener("keydown", e => { if (e.key === "Enter") _renderizarTabelaColab(); });
    // Anonimizar
    document.getElementById("switchAnonimizar").addEventListener("change", e => {
      S.anonimizar = e.target.checked;
      window.AF_ANON = S.anonimizar;
      if (S.liquido) { _renderizarTabela(); _renderOutliers(); }
    });
    // Reset dados
    document.getElementById("btnResetDados").addEventListener("click", acaoResetDados);
    // Visão Geral: busca por colaborador/matrícula no painel de descontos
    const vgB = document.getElementById("vgBusca");
    if (vgB) vgB.addEventListener("input", e => { S.vgBusca = e.target.value; if (S.descontosRevisar) _renderPainelDescontos(); });
    const vgBL = document.getElementById("vgBuscaLimpar");
    if (vgBL) vgBL.addEventListener("click", () => { S.vgBusca = ""; const i = document.getElementById("vgBusca"); if (i) i.value = ""; if (S.descontosRevisar) _renderPainelDescontos(); });

  }

  function _popularComboMetodo() {
    const sel = document.getElementById("comboMetodo");
    for (const nome of Object.keys(PARAMETROS_MERCADO)) {
      const o = document.createElement("option"); o.value = nome; o.textContent = nome; sel.appendChild(o);
    }
    sel.value = S.configNome;
    document.getElementById("lblMetodoDesc").textContent = PARAMETROS_MERCADO[S.configNome].descricao;
  }

  function _switchTab(tabId) {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === tabId));
    document.querySelectorAll(".tab-pane").forEach(p => p.classList.toggle("active", p.id === `tab-${tabId}`));
    // Lazy render: charts em aba inativa têm canvas com dimensão zero
    if (S.liquido && S.tabsDirty[tabId]) {
      requestAnimationFrame(() => {
        if (tabId === "outliers")      _renderOutliers();
        if (tabId === "analiseVerbas") _renderAnaliseVerbas();
        delete S.tabsDirty[tabId];
      });
    }
    if (tabId === "conferencia" && S.liquido) _renderConferencia();
    if (tabId === "holerite" && S.liquido) _renderHolerite();
  }

  // ---- File handling ----
  function _onFileChange(e) {
    const file = e.target.files[0]; if (!file) return;
    document.getElementById("fileName").textContent = file.name;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const csvText = ev.target.result;
        const { df, meses } = carregarCSV(csvText);
        if (!df.length || !meses.length) { alert("Arquivo sem dados reconhecíveis. Verifique o formato CSV (separador ; e valores com vírgula decimal)."); return; }
        RowSelection.newDataset();
        df.forEach(RowSelection.sourceId);
        S.df = df; S.mesesDetectados = meses;
        S.csvBruto = csvText; S.csvNome = file.name;
        _salvarNoStorage(csvText, file.name);
        _carregarEmpresas(); _carregarProcessos();
        _popularComboMesAlvo(); _renderBaselineCheckboxes();
        ["btnReprocessar","btnExportar","btnFiltroVerbas"].forEach(id => document.getElementById(id).disabled = false);
        const si = document.getElementById("stateInicial"); if (si) si.style.display = "none";
        const vc = document.getElementById("vgContent"); if (vc) vc.style.display = "";
        acaoReprocessar();
      } catch (err) { alert(`Erro ao ler CSV: ${err.message}`); console.error(err); }
    };
    reader.readAsText(file, "ISO-8859-1");
    e.target.value = "";
  }

  // ---- Persistência localStorage ----
  function _salvarNoStorage(csvText, fname) {
    try {
      localStorage.setItem("af_csv", csvText);
      localStorage.setItem("af_nome", fname);
      _atualizarCache(`Em cache: ${fname}`);
    } catch (e) {
      _atualizarCache("Arquivo grande demais para cache local.");
    }
  }

  function _atualizarCache(msg) {
    const el = document.getElementById("lblCache"); if (el) el.textContent = msg;
    const btn = document.getElementById("btnResetDados");
    if (btn) btn.style.display = msg ? "" : "none";
  }

  function _loadFromStorage() {
    const csvText = localStorage.getItem("af_csv");
    const fname   = localStorage.getItem("af_nome");
    if (!csvText) return;
    try {
      const { df, meses } = carregarCSV(csvText);
      if (!df.length || !meses.length) { localStorage.removeItem("af_csv"); return; }
      RowSelection.newDataset();
        df.forEach(RowSelection.sourceId);
        S.df = df; S.mesesDetectados = meses; S.csvBruto = csvText; S.csvNome = fname;
      document.getElementById("fileName").textContent = fname || "(cache)";
      _atualizarCache(`Em cache: ${fname || "arquivo anterior"}`);
      _carregarEmpresas(); _carregarProcessos();
      _popularComboMesAlvo(); _renderBaselineCheckboxes();
      ["btnReprocessar","btnExportar","btnFiltroVerbas"].forEach(id => document.getElementById(id).disabled = false);
      const si = document.getElementById("stateInicial"); if (si) si.style.display = "none";
      const vc = document.getElementById("vgContent"); if (vc) vc.style.display = "";
      acaoReprocessar();
    } catch (e) { localStorage.removeItem("af_csv"); console.warn("Cache inválido, ignorado."); }
  }

  function acaoResetDados() {
    if (!confirm("Limpar os dados salvos e reiniciar o app?")) return;
    localStorage.removeItem("af_csv");
    localStorage.removeItem("af_nome");
    location.reload();
  }

  // ============================================================
  // ---- Conferência (revisão humana) — portado do Recibo ------
  // ============================================================
  const CONF_KEY = "af_conf_v2"; // v2: chave agora inclui a competência (mês-alvo)
  function _confComp() { return String(S.mesAlvo || ""); }
  // Conferência por competência: "conferido" de um mês não reaparece no mês seguinte.
  // (Dimensão Empresa na conferência fica como próximo passo — exige empresa no contexto do holerite.)
  function _confKey(r)   { return String(r.Matrícula) + "|||" + String(r.Nome) + "|||" + _confComp(); }
  function _confKeyMN(m, n) { return String(m) + "|||" + String(n) + "|||" + _confComp(); }
  function _lineKey(row) { return String(row["Código"]||"") + "||" + String(row["Descrição"]||"") + "||" + String(row["Clas."]||""); }
  function _confLoad() {
    try { const r = JSON.parse(localStorage.getItem(CONF_KEY)); S.conf = (r && typeof r === "object") ? r : {}; }
    catch (e) { S.conf = {}; }
  }
  function _confSave() { try { localStorage.setItem(CONF_KEY, JSON.stringify(S.conf)); } catch (e) {} }
  function _confGet(key) { return S.conf[key] || { ok: false, obs: "", linhas: {} }; }
  function _confEnsure(key) {
    if (!S.conf[key]) S.conf[key] = { ok: false, obs: "", linhas: {} };
    if (!S.conf[key].linhas) S.conf[key].linhas = {};
    return S.conf[key];
  }
  function _confIsEmpty(c) { return !c.ok && !(c.obs && c.obs.trim()) && !(c.linhas && Object.keys(c.linhas).length); }
  function _confSetFunc(key, val) {
    const c = _confEnsure(key); c.ok = !!val; c.ts = new Date().toISOString();
    if (_confIsEmpty(c)) delete S.conf[key];
    _confSave(); _atualizarConfContador();
  }
  function _confSetObs(key, txt) {
    const c = _confEnsure(key); c.obs = txt; c.ts = new Date().toISOString();
    if (_confIsEmpty(c)) delete S.conf[key];
    _confSave();
  }
  function _confSetLinha(key, lk, val) {
    const c = _confEnsure(key);
    if (val) c.linhas[lk] = true; else delete c.linhas[lk];
    c.ts = new Date().toISOString();
    if (_confIsEmpty(c)) delete S.conf[key];
    _confSave();
  }
  function _confStats() {
    const pop = S.liquidoAnalise || [];
    let ok = 0, comObs = 0;
    pop.forEach(r => { const c = S.conf[_confKey(r)]; if (c && c.ok) ok++; if (c && c.obs && c.obs.trim()) comObs++; });
    return { ok, total: pop.length, pend: pop.length - ok, comObs };
  }
  function _atualizarConfContador() {
    const s = _confStats();
    const pct = s.total ? Math.round(s.ok / s.total * 100) : 0;
    const txt = `Conferidos: ${s.ok} de ${s.total} (${pct}%) · Pendentes: ${s.pend} · Com observação: ${s.comObs}`;
    const el = document.getElementById("lblConfContador");
    if (el) el.textContent = txt;
    const elt = document.getElementById("lblConfContadorTab");
    if (elt) elt.textContent = txt;
    const fill = document.getElementById("confProgFill");
    if (fill) fill.style.width = pct + "%";
    const ex = document.getElementById("confExpStatus");
    if (ex && ex.dataset.live === "1") ex.textContent = `${s.ok} conferidos · ${s.pend} pendentes · ${s.comObs} com observação.`;
  }
  // Célula de checkbox para tabelas de funcionário (delegação trata o change)
  function _confCellFunc(r) {
    const key = _confKey(r);
    const c = _confGet(key);
    const nLin = c.linhas ? Object.keys(c.linhas).length : 0;
    const dot = (c.obs && c.obs.trim()) ? `<span class="conf-obs-dot" title="${escAttr(c.obs)}">obs</span>` : "";
    const lin = nLin ? `<span class="conf-lin-badge" title="${nLin} verba(s) conferida(s)">${nLin}</span>` : "";
    return `<td class="center conf-cell" data-confkey="${escAttr(key)}">
      <input type="checkbox" class="conf-chk" onclick="event.stopPropagation()" ${c.ok ? "checked" : ""} title="Marcar funcionário como conferido">${dot}${lin}</td>`;
  }
  function escAttr(s) { return String(s == null ? "" : s).replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function escHtml(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  // Delegação de eventos para checkboxes de funcionário em qualquer tbody
  function _bindConfDelegation(tbodyId) {
    const tb = document.getElementById(tbodyId);
    if (!tb || tb.dataset.confBound === "1") return;
    tb.dataset.confBound = "1";
    tb.addEventListener("change", e => {
      const chk = e.target.closest(".conf-chk"); if (!chk) return;
      const cell = chk.closest("[data-confkey]"); if (!cell) return;
      _confSetFunc(cell.dataset.confkey, chk.checked);
      // se o filtro de conferência estiver ativo, re-renderiza para refletir
      if (S.filtroConf !== "todos") { _renderizarTabela(); }
      // sincroniza caixa de drill se for o selecionado
      if (S.funcSelecionado && _confKey(S.funcSelecionado) === cell.dataset.confkey) _renderConfBox(S.funcSelecionado);
    });
  }
  // Caixa de conferência no drill do funcionário
  function _renderConfBox(r) {
    const box = document.getElementById("confBox"); if (!box) return;
    if (!r) { box.style.display = "none"; return; }
    box.style.display = "";
    const key = _confKey(r); const c = _confGet(key);
    document.getElementById("confBoxChk").checked = c.ok;
    document.getElementById("confBoxObs").value = c.obs || "";
    const nLin = c.linhas ? Object.keys(c.linhas).length : 0;
    document.getElementById("confBoxInfo").textContent =
      `${c.ok ? "Conferido" : "Pendente"}${nLin ? " · " + nLin + " verba(s) marcada(s)" : ""}${c.ts ? " · " + new Date(c.ts).toLocaleString("pt-BR") : ""}`;
  }
  function _bindConfBox() {
    const chk = document.getElementById("confBoxChk");
    const obs = document.getElementById("confBoxObs");
    if (chk && !chk.dataset.b) { chk.dataset.b = "1";
      chk.addEventListener("change", () => {
        if (!S.funcSelecionado) return;
        _confSetFunc(_confKey(S.funcSelecionado), chk.checked);
        _renderConfBox(S.funcSelecionado); _renderizarTabela(); _repintarColabConf();
      });
    }
    if (obs && !obs.dataset.b) { obs.dataset.b = "1";
      obs.addEventListener("input", () => {
        if (!S.funcSelecionado) return;
        _confSetObs(_confKey(S.funcSelecionado), obs.value);
        _atualizarConfContador();
      });
    }
  }
  // Repinta checkboxes da tabela colab sem re-render completo
  function _repintarColabConf() {
    document.querySelectorAll("#colabBody [data-confkey]").forEach(cell => {
      const c = _confGet(cell.dataset.confkey);
      const chk = cell.querySelector(".conf-chk"); if (chk) chk.checked = c.ok;
    });
  }
  // Checkbox por verba no drill
  function _confCellLinha(r, row) {
    const key = _confKey(r); const lk = _lineKey(row);
    const on = !!(_confGet(key).linhas || {})[lk];
    return `<td class="center conf-cell-lin" data-confkey="${escAttr(key)}" data-linekey="${escAttr(lk)}">
      <input type="checkbox" class="conf-chk-lin" ${on ? "checked" : ""} title="Marcar esta verba como conferida"></td>`;
  }
  function _bindConfLinhaDelegation() {
    const tb = document.getElementById("drillBody");
    if (!tb || tb.dataset.confLinBound === "1") return;
    tb.dataset.confLinBound = "1";
    tb.addEventListener("change", e => {
      const chk = e.target.closest(".conf-chk-lin"); if (!chk) return;
      const cell = chk.closest("[data-linekey]"); if (!cell) return;
      _confSetLinha(cell.dataset.confkey, cell.dataset.linekey, chk.checked);
      if (S.funcSelecionado) _renderConfBox(S.funcSelecionado);
    });
  }
  // ---- Export / Import / Reset da conferência ----
  function confExport() {
    const blob = new Blob([JSON.stringify({
      version: 1, app: "Auditoria de Folha", exportedAt: new Date().toISOString(),
      mesAlvo: S.mesAlvo, conf: S.conf
    }, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `conferencia_${(S.mesAlvo || "").replace("/", "-")}.json`;
    a.click(); URL.revokeObjectURL(a.href);
    const el = document.getElementById("confExpStatus");
    if (el) { el.dataset.live = "1"; _atualizarConfContador(); }
  }
  function confImport(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const d = JSON.parse(e.target.result);
        const incoming = d && d.conf ? d.conf : (d && typeof d === "object" ? d : null);
        if (!incoming) { alert("Arquivo de conferência inválido."); return; }
        const modo = confirm("OK = mesclar com a conferência atual.\nCancelar = substituir tudo pela do arquivo.");
        if (modo) { Object.keys(incoming).forEach(k => { S.conf[k] = incoming[k]; }); }
        else { S.conf = incoming; }
        _confSave(); _atualizarConfContador(); _renderizarTabela();
        if (S.funcSelecionado) _renderConfBox(S.funcSelecionado);
        _renderConferencia();
        alert("Conferência importada.");
      } catch (err) { alert("Falha ao ler o JSON de conferência."); }
    };
    reader.readAsText(file);
  }
  function confReset() {
    if (!confirm("Limpar TODAS as marcações de conferência e observações? Esta ação não pode ser desfeita.")) return;
    S.conf = {}; _confSave(); _atualizarConfContador(); _renderizarTabela();
    if (S.funcSelecionado) _renderConfBox(S.funcSelecionado);
    _renderConferencia();
  }

  // ---- Aba Conferência (visão geral consolidada) ----
  function _renderConferencia() {
    const tbody = document.getElementById("confBody"); if (!tbody) return;
    tbody.innerHTML = "";
    if (!S.liquidoAnalise) { document.getElementById("countConf").textContent = ""; _atualizarConfContador(); return; }
    let list = [...S.liquidoAnalise];
    const f = document.getElementById("comboConfFiltro2")?.value || "todos";
    const busca = (document.getElementById("entryConfBusca2")?.value || "").toUpperCase().trim();
    if (busca) list = list.filter(r => r.Nome.toUpperCase().includes(busca) || String(r.Matrícula).includes(busca));
    if (f === "pendentes")       list = list.filter(r => !(S.conf[_confKey(r)] && S.conf[_confKey(r)].ok));
    else if (f === "conferidos") list = list.filter(r => S.conf[_confKey(r)] && S.conf[_confKey(r)].ok);
    else if (f === "comobs")     list = list.filter(r => { const c = S.conf[_confKey(r)]; return c && c.obs && c.obs.trim(); });
    else if (f === "criticos")   list = list.filter(r => STATUS_CRITICOS.has(r.STATUS));
    // Pendentes primeiro; dentro de cada grupo, maior impacto — ou ordenação por clique
    if (S.confSort.key) {
      list.sort(_genCmp(_confValSort, S.confSort.key, S.confSort.dir));
    } else {
      list.sort((a, b) => {
        const ca = (S.conf[_confKey(a)] && S.conf[_confKey(a)].ok) ? 1 : 0;
        const cb = (S.conf[_confKey(b)] && S.conf[_confKey(b)].ok) ? 1 : 0;
        if (ca !== cb) return ca - cb;
        return Math.abs(b.VAR_ABS || 0) - Math.abs(a.VAR_ABS || 0);
      });
    }
    const sum = _sumView('conferencia', list, {render: _renderConferencia, investigate: r => _navigateToFuncDrill(r.Matrícula, r.Nome, r.Empresa)});
    sum.page.forEach(r => {
      const key = _confKey(r); const c = _confGet(key);
      const nLin = c.linhas ? Object.keys(c.linhas).length : 0;
      const tr = document.createElement("tr");
      if (c.ok) tr.className = "conf-row-ok";
      else if (STATUS_CRITICOS.has(r.STATUS)) tr.className = "critico";
      const dNome = maskNome(r.Nome, r.Matrícula), dMat = maskMat(r.Matrícula);
      tr.innerHTML =
        `<td class="center conf-cell" data-confkey="${escAttr(key)}"><input type="checkbox" class="conf-chk" ${c.ok ? "checked" : ""} title="Marcar como conferido"></td>
        <td class="center">${dMat}</td>
        <td>${String(dNome).substring(0,55)}</td>
        <td class="center"><span class="sbadge ${_badgeCls(r.STATUS)}">${r.STATUS}</span></td>
        <td class="num">${fmtBRL(r.LIQUIDO_ALVO)}</td>
        <td class="num">${fmtBRL(r.VAR_ABS)}</td>
        <td class="center">${nLin || ""}</td>
        <td><input type="text" class="conf-obs-inp" data-confkey="${escAttr(key)}" value="${escAttr(c.obs || "")}" placeholder="observação..."></td>
        <td class="center"><button class="btn-xs conf-open" data-empresa="${escAttr(r.Empresa)}" data-mat="${escAttr(r.Matrícula)}" data-nome="${escAttr(r.Nome)}">Abrir</button></td>`;
      sum.view.bind(tr, r); tbody.appendChild(tr);
    });
    sum.view.finish(document.getElementById("tabelaConf"));
    _bindConfTab();
    _bindSort("#tabelaConf", S.confSort, _renderConferencia);
    document.getElementById("countConf").textContent = `${sum.page.length} de ${list.length} colaboradores`;
    _atualizarConfContador();
  }
  function _confValSort(r, key) {
    const c = S.conf[_confKey(r)] || {};
    switch (key) {
      case "conf":   return c.ok ? 1 : 0;
      case "mat":    return String(r.Matrícula);
      case "nome":   return String(r.Nome);
      case "status": return String(r.STATUS);
      case "liq":    return r.LIQUIDO_ALVO || 0;
      case "var":    return Math.abs(r.VAR_ABS || 0);
      case "nlin":   return c.linhas ? Object.keys(c.linhas).length : 0;
      case "obs":    return (c.obs || "").trim() ? 1 : 0;
      default:       return "";
    }
  }
  function _bindConfTab() {
    const tb = document.getElementById("confBody");
    if (!tb || tb.dataset.b === "1") return; tb.dataset.b = "1";
    tb.addEventListener("change", e => {
      const chk = e.target.closest(".conf-chk"); if (!chk) return;
      const cell = chk.closest("[data-confkey]"); if (!cell) return;
      _confSetFunc(cell.dataset.confkey, chk.checked);
      _renderizarTabela(); _repintarColabConf();
      if (S.funcSelecionado && _confKey(S.funcSelecionado) === cell.dataset.confkey) _renderConfBox(S.funcSelecionado);
      const tr = chk.closest("tr"); if (tr) tr.classList.toggle("conf-row-ok", chk.checked);
      const f = document.getElementById("comboConfFiltro2").value;
      if (f === "pendentes" || f === "conferidos") _renderConferencia();
    });
    tb.addEventListener("input", e => {
      const inp = e.target.closest(".conf-obs-inp"); if (!inp) return;
      _confSetObs(inp.dataset.confkey, inp.value);
      _atualizarConfContador();
    });
    tb.addEventListener("click", e => {
      const btn = e.target.closest(".conf-open"); if (!btn) return;
      _navigateToFuncDrill(btn.dataset.mat, btn.dataset.nome, btn.dataset.empresa);
    });
  }
  function _confMarcarVisiveis() {
    const cells = document.querySelectorAll("#confBody td.conf-cell[data-confkey]");
    if (!cells.length) return;
    if (!confirm(`Marcar ${cells.length} colaborador(es) visível(eis) como conferido(s)?`)) return;
    cells.forEach(cell => _confSetFunc(cell.dataset.confkey, true));
    _renderConferencia(); _renderizarTabela(); _repintarColabConf();
  }

  // ============================================================
  // ---- Holerite (recibo do colaborador) — aba secundária -----
  // ============================================================
  let HOL = { empresa: null, mat: null, nome: null, q: "", onlyChk: false, buscaColab: "", sort: { key: null, dir: 1 },
    clas: { pgto: true, desc: true, prov: true, outro: true, base: true, encar: true } };

  // Ordenação genérica por clique no cabeçalho (asc/desc)
  function _genCmp(valFn, key, dir) {
    return (a, b) => {
      const va = valFn(a, key), vb = valFn(b, key);
      if (typeof va === "number" && typeof vb === "number") return dir * (va - vb);
      return dir * String(va == null ? "" : va).localeCompare(String(vb == null ? "" : vb), "pt-BR", { numeric: true });
    };
  }
  function _bindSort(sel, state, rerender) {
    document.querySelectorAll(sel + " th[data-sk]").forEach(th => {
      th.classList.add("sortable"); th.style.cursor = "pointer";
      th.classList.remove("sort-asc", "sort-desc");
      if (state.key === th.dataset.sk) th.classList.add(state.dir > 0 ? "sort-asc" : "sort-desc");
      th.onclick = () => {
        const k = th.dataset.sk;
        if (state.key === k) state.dir *= -1; else { state.key = k; state.dir = 1; }
        rerender();
      };
    });
  }

  function _clasKind(c) {
    const n = String(c == null ? "" : c).toLowerCase();
    if (n.indexOf("pgto") > -1 || n.indexOf("pgt") > -1 || n.indexOf("pagam") > -1) return "pgto";
    if (n.indexOf("outro") > -1) return "outro";
    if (n.indexOf("desc") > -1) return "desc";
    if (n.indexOf("base") > -1) return "base";
    if (n.indexOf("prov") > -1) return "prov";
    if (n.indexOf("encar") > -1) return "encar";
    return "outro";
  }
  function _clasRank(c) { return ({ pgto: 1, outro: 3, desc: 2, base: 5, prov: 4, encar: 6 })[_clasKind(c)] || 99; }
  function _isLiqRow(row) {
    if (String(row["Código"] || "").replace(/^0+/, "") === "9950") return true;
    return String((row["Descrição"] || "") + " " + (row["Código"] || "")).toLowerCase().indexOf("liquido") > -1
        || String(row["Descrição"] || "").toLowerCase().indexOf("líquido") > -1;
  }
  function _linReg(seq) {
    const n = seq.length; let sx = 0, sy = 0, sxx = 0, sxy = 0;
    for (let i = 0; i < n; i++) { sx += i; sy += seq[i]; sxx += i * i; sxy += i * seq[i]; }
    const d = n * sxx - sx * sx;
    const slope = d ? (n * sxy - sx * sy) / d : 0;
    return { slope, intercept: (sy - slope * sx) / n };
  }
  function _sparkSvg(seq) {
    const w = 96, h = 28, pad = 4, n = seq.length;
    if (!n) return "";
    const mn = Math.min(...seq), mx = Math.max(...seq), rng = (mx - mn) || 1;
    const x = i => n === 1 ? w / 2 : pad + i * (w - 2 * pad) / (n - 1);
    const y = v => { const yy = h - pad - ((v - mn) / rng) * (h - 2 * pad); return Math.max(pad, Math.min(h - pad, yy)); };
    const first = seq[0], last = seq[n - 1];
    const { slope, intercept } = _linReg(seq);
    const scaleRef = Math.max(Math.abs(mx), Math.abs(mn), 1);
    const flat = rng / scaleRef < 0.005;
    const dir = flat ? 0 : (slope > 0 ? 1 : (slope < 0 ? -1 : 0));
    const col = dir > 0 ? "#005A64" : (dir < 0 ? "#8C321E" : "#646464");
    const pctTot = Math.abs(first) > 0.005 ? (last - first) / Math.abs(first) * 100 : (Math.abs(last) > 0.005 ? 100 * Math.sign(last) : 0);
    const resid = seq.map((v, i) => v - (intercept + slope * i));
    const sd = Math.sqrt(resid.reduce((a, b) => a + b * b, 0) / n);
    const outIdx = [];
    if (n >= 4 && sd > 0) { const thr = Math.max(1.6 * sd, 0.02 * scaleRef); seq.forEach((v, i) => { if (Math.abs(resid[i]) > thr) outIdx.push(i); }); }
    const sign = pctTot > 0 ? "+" : "";
    const dirTxt = dir > 0 ? "alta" : (dir < 0 ? "queda" : "estável");
    const title = `Tendência ${dirTxt} · variação total ${sign}${pctTot.toFixed(1)}%${outIdx.length ? " · " + outIdx.length + " ponto(s) fora da tendência" : ""}`;
    if (n === 1) return `<svg class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><title>${title}</title><circle cx="${x(0).toFixed(1)}" cy="${y(seq[0]).toFixed(1)}" r="2.4" fill="${col}"/></svg>`;
    const pts = seq.map((v, i) => `${x(i).toFixed(1)},${y(v).toFixed(1)}`).join(" ");
    const area = `${x(0).toFixed(1)},${(h - pad).toFixed(1)} ${pts} ${x(n - 1).toFixed(1)},${(h - pad).toFixed(1)}`;
    const ty0 = y(intercept), ty1 = y(intercept + slope * (n - 1));
    let dots = "";
    seq.forEach((v, i) => {
      const isOut = outIdx.indexOf(i) >= 0, isLast = i === n - 1;
      if (isOut) dots += `<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="3.1" fill="#7D0041" stroke="#fff" stroke-width="1"/>`;
      else if (isLast) dots += `<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="2.3" fill="${col}"/>`;
      else dots += `<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="1.5" fill="${col}" fill-opacity=".75"/>`;
    });
    return `<svg class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><title>${title}</title>`
      + `<polyline fill="${col}" fill-opacity=".10" stroke="none" points="${area}"/>`
      + `<line x1="${x(0).toFixed(1)}" y1="${ty0.toFixed(1)}" x2="${x(n - 1).toFixed(1)}" y2="${ty1.toFixed(1)}" stroke="${col}" stroke-width="1" stroke-dasharray="3 2" stroke-opacity=".55"/>`
      + `<polyline fill="none" stroke="${col}" stroke-width="1.6" stroke-linejoin="round" stroke-linecap="round" points="${pts}"/>${dots}</svg>`;
  }
  // Linhas do holerite para o colaborador atual
  function _holLines() {
    if (!HOL.mat || !S.dfFiltrado) return [];
    const rows = S.dfFiltrado.filter(row => getMatricula(row) === HOL.mat && getNome(row) === HOL.nome && getEmpresa(row) === HOL.empresa);
    return rows.map(row => ({
      codigo: row["Código"], descricao: row["Descrição"], clas: row["Clas."],
      cr: row["C.R."], processo: row["Processo"], row,
      liq: _isLiqRow(row),
      lk: _lineKey(row),
      per: Object.fromEntries(S.mesesDetectados.map(m => [m, row[`${m} - Valor`] || 0]))
    }));
  }
  function _holLineChecked(l) {
    const c = S.conf[_confKeyMN(HOL.mat, HOL.nome)];
    return !!(c && c.linhas && c.linhas[l.lk]);
  }
  function _holLiqValue(lines) {
    const l = lines.find(x => x.liq);
    return l ? { found: true, value: l.per[S.mesAlvo] || 0 } : { found: false, value: 0 };
  }
  function _holCalc(lines) {
    let creditos = 0, descontos = 0, outros = 0, count = 0;
    lines.forEach(l => {
      if (l.liq || !_holLineChecked(l)) return;
      const v = l.per[S.mesAlvo] || 0, k = _clasKind(l.clas);
      count++;
      if (k === "desc") descontos += v;
      else if (k === "pgto" || k === "prov") creditos += v;
      else outros += v;
    });
    return { count, creditos, descontos, outros, calculado: creditos + outros - descontos };
  }
  function _holClasCounts(lines) {
    const p = {}; lines.forEach(l => { if (l.liq) return; const k = _clasKind(l.clas); p[k] = (p[k] || 0) + 1; });
    return p;
  }
  function _popularHolFunc() {
    const sel = document.getElementById("holFunc"); if (!sel) return;
    const q = (HOL.buscaColab || "").toLowerCase().trim();
    let pop = (S.liquidoAnalise || []).slice().sort((a, b) => String(a.Nome).localeCompare(String(b.Nome)));
    if (q) pop = pop.filter(r => String(r.Nome).toLowerCase().includes(q) || String(r.Matrícula).toLowerCase().includes(q));
    sel.innerHTML = pop.map(r => `<option value="${escAttr(JSON.stringify([r.Empresa, r.Matrícula, r.Nome]))}">${escAttr(maskMat(r.Matrícula))} — ${escAttr(String(maskNome(r.Nome, r.Matrícula)).substring(0,45))} (${r.STATUS})</option>`).join("");
    const cont = document.getElementById("holCount");
    if (cont) cont.textContent = q ? `${pop.length} colaborador(es) encontrado(s)` : "";
    if (!pop.length) { HOL.mat = null; HOL.nome = null; return; }
    const cur = HOL.mat ? JSON.stringify([HOL.empresa, HOL.mat, HOL.nome]) : null;
    const has = cur && pop.some(r => (JSON.stringify([r.Empresa, r.Matrícula, r.Nome])) === cur);
    let alvo = has ? cur
      : (!q && S.funcSelecionado && pop.some(r => r.Matrícula === S.funcSelecionado.Matrícula && r.Nome === S.funcSelecionado.Nome)
          ? JSON.stringify([S.funcSelecionado.Empresa, S.funcSelecionado.Matrícula, S.funcSelecionado.Nome])
          : JSON.stringify([pop[0].Empresa, pop[0].Matrícula, pop[0].Nome]));
    sel.value = alvo; const [emp, m, n] = JSON.parse(alvo); HOL.empresa = emp; HOL.mat = m; HOL.nome = n;
  }
  function _renderHolerite() {
    const cont = document.getElementById("holContent"); if (!cont) return;
    if (!S.liquido) { cont.innerHTML = `<div class="hint-text">Carregue um arquivo e processe para ver o holerite.</div>`; return; }
    _popularHolFunc();
    if (!HOL.mat) {
      const v = RowSelection.table('holerite'); v.prepare([], { key: l => RowSelection.sourceId(l), values: () => ({}), render: _renderHolerite });
      const notice = v.message; cont.innerHTML = `<div class="hint-text">Selecione um colaborador.</div>`;
      if (notice) { const el = document.createElement('div'); el.setAttribute('role', 'status'); el.textContent = notice; cont.appendChild(el); }
      return;
    }
    const fr = (S.liquido || []).find(r => r.Matrícula === HOL.mat && r.Nome === HOL.nome && r.Empresa === HOL.empresa) || {};
    let lines = _holLines();
    const totalLines = lines.length;
    // ordena: líquido primeiro, depois por classificação
    lines.sort((a, b) => (a.liq ? -1 : b.liq ? 1 : _clasRank(a.clas) - _clasRank(b.clas)) || String(a.codigo).localeCompare(String(b.codigo)));
    // filtro
    const q = HOL.q.toLowerCase().trim();
    let visiveis = lines.filter(l => {
      if (l.liq) return true;
      if (!HOL.clas[_clasKind(l.clas)]) return false;
      if (HOL.onlyChk && !_holLineChecked(l)) return false;
      if (q) { const hay = String((l.codigo || "") + " " + (l.descricao || "") + " " + (l.clas || "")).toLowerCase(); if (hay.indexOf(q) < 0) return false; }
      return true;
    });
    const hiddenN = totalLines - visiveis.length;
    // ordenação por clique (líquido fica fixo no topo como referência)
    if (HOL.sort.key) {
      const liqLine = visiveis.find(l => l.liq);
      const rest = visiveis.filter(l => !l.liq).sort(_genCmp(_holValSort, HOL.sort.key, HOL.sort.dir));
      visiveis = (liqLine ? [liqLine] : []).concat(rest);
    }
    const okN = lines.filter(l => !l.liq && _holLineChecked(l)).length;
    const calcEff = totalLines - 1; // exclui líquido da contagem-base
    const liq = _holLiqValue(lines);
    if (!liq.found && fr && fr.LIQUIDO_ALVO != null) { liq.found = true; liq.value = fr.LIQUIDO_ALVO; }
    const c = _holCalc(lines);

    const counts = _holClasCounts(lines);
    const CHIPS = [["pgto","PGTO"],["desc","DESC"],["prov","PROV"],["outro","OUTRO"],["base","BASE"],["encar","ENCAR"]];
    const chips = CHIPS.filter(([k]) => counts[k]).map(([k, lbl]) =>
      `<button class="hol-chip tag ${k} ${HOL.clas[k] ? "on" : "off"}" data-clas="${k}">${lbl} <span class="n">${counts[k]}</span></button>`).join("");

    // cards de reconciliação
    let dif = 0, cls = "warn", status = "Selecione as verbas de pagamento e desconto para validar.";
    if (!liq.found) { status = "Não localizei a verba Líquido (9950) neste holerite."; }
    else if (c.count > 0) { dif = c.calculado - liq.value;
      if (Math.abs(dif) < 0.01) { status = "CERTO: os selecionados fecham com o Líquido."; cls = "ok"; }
      else { status = "DIVERGENTE: diferença entre a seleção e o Líquido."; cls = "err"; } }

    const months = S.mesesDetectados;
    const head = `<tr>
      <th class="hol-sel" data-sk="calc">Calc.</th><th data-sk="cod">Cód.</th><th data-sk="desc">Descrição</th><th class="c" data-sk="clas">Clas.</th>
      ${months.map(m => `<th class="r ${m === S.mesAlvo ? "hol-alvo" : ""}" data-sk="m:${m}">${m}</th>`).join("")}
      <th class="c hol-sparkhead" data-sk="trend">Tendência</th></tr>`;
    const sum = _sumView('holerite', visiveis, {
      key: l => JSON.stringify([RowSelection.sourceId(l.row), S.mesAlvo]), values: l => _sumPayroll(l.row), emptyGroup: 'Folha (proventos − descontos)',
      render: _holReRender, investigate: l => RowSelection.openDetail('Lançamento de folha · ' + S.mesAlvo, l.row)
    });
    const body = sum.page.map(l => {
      const seq = months.map(m => l.per[m] || 0);
      const kind = _clasKind(l.clas);
      const checked = _holLineChecked(l);
      const cells = months.map((m, i) => {
        const v = seq[i]; let hm = "", arw = "";
        if (i > 0) { const pr = seq[i - 1];
          if (pr === 0 && Math.abs(v) > 0.005) { hm = "hm-new"; arw = ' <span class="arw up">▲</span>'; }
          else if (Math.abs(pr) > 0.005 && Math.abs(v) < 0.005) { hm = "hm-ext"; arw = ' <span class="arw dn">▼</span>'; }
          else if (Math.abs(pr) > 0.005) { const pct = (v - pr) / Math.abs(pr) * 100, a = Math.abs(pct);
            if (a >= 98) hm = "hm-ext"; else if (a >= 30) hm = "hm-not";
            if (a >= 30) arw = ` <span class="arw ${pct > 0 ? "up" : "dn"}">${pct > 0 ? "▲" : "▼"}</span>`; } }
        return `<td class="r mcell ${hm} ${m === S.mesAlvo ? "hol-alvo" : ""}">${fmtBRL(v)}${arw}</td>`;
      }).join("");
      return `<tr class="${checked ? "line-ok" : ""} ${l.liq ? "liquido" : ""}">
        <td class="hol-sel">${l.liq ? "" : `<input type="checkbox" class="hol-check" data-lk="${escAttr(l.lk)}" ${checked ? "checked" : ""} title="Incluir/remover no cálculo">`}</td>
        <td class="code-n">${escAttr(l.codigo)}</td>
        <td title="${escAttr(l.descricao)}">${String(l.descricao || "").substring(0,46)}${l.liq ? ' <span class="tag pgto">LÍQ</span>' : ""}</td>
        <td class="c"><span class="tag ${kind}">${escAttr(l.clas || "-")}</span></td>
        ${cells}
        <td class="c hol-sparkcell">${_sparkSvg(seq)}</td></tr>`;
    }).join("");

    cont.innerHTML = `
      <div class="hol-head">
        <div><div class="hol-nome">${escAttr(maskNome(fr.Nome, fr.Matrícula))}</div>
          <div class="hol-sub">Matrícula ${escAttr(maskMat(fr.Matrícula))} · ${escAttr((fr.Empresa || "").trim() || "-")} · Processo ${escAttr(fr.Processo || "-")}</div></div>
        <div class="hol-meta">Status <span class="sbadge ${_badgeCls(fr.STATUS)}">${fr.STATUS || "-"}</span><br>
          Mês base <b>${S.mesAlvo}</b> · ${totalLines} lançamento(s)</div>
      </div>
      <div class="hol-filterbar">
        <input type="text" id="holBusca" class="input-ctrl-sm" placeholder="Filtrar verba por código ou descrição..." value="${escAttr(HOL.q)}" style="min-width:240px">
        <span class="hol-chips">${chips}</span>
        <label class="conf-box-chk" style="font-size:12px"><input type="checkbox" id="holOnlyChk" ${HOL.onlyChk ? "checked" : ""}> Só marcadas (Calc.)</label>
      </div>
      <div class="calc-box">
        <div class="calc-card"><div class="l">Líquido localizado</div><div class="v">${liq.found ? fmtBRL(liq.value) : "-"}</div></div>
        <div class="calc-card"><div class="l">Créditos marcados (Calc.)</div><div class="v">${fmtBRL(c.creditos + c.outros)}</div></div>
        <div class="calc-card"><div class="l">Descontos marcados (Calc.)</div><div class="v">${fmtBRL(c.descontos)}</div></div>
        <div class="calc-card ${cls}"><div class="l">Resultado</div><div class="v">${fmtBRL(c.calculado)}</div></div>
        <div class="calc-card ${cls}"><div class="l">Diferença</div><div class="v">${liq.found && c.count > 0 ? fmtBRL(dif) : "-"}</div></div>
        <div class="calc-note"><b>${status}</b><br>Marque na coluna <b>Calc.</b> as linhas que devem compor a conta. Regra: PGTO/PROV/OUTRO somam; DESC subtrai; a linha Líquido é só referência.</div>
      </div>
      <div class="hol-linecount">Linhas marcadas em Calc. (conferência): <b>${okN}</b> de <b>${calcEff}</b> (${calcEff ? Math.round(okN / calcEff * 100) : 0}%).${hiddenN > 0 ? ` · <b>${hiddenN}</b> verba(s) ocultada(s) pelos filtros — a verba Líquido permanece sempre visível.` : ""}</div>
      <div class="hol-legend"><b>${months.length}</b> meses na tela (base <b>${S.mesAlvo}</b>). <span class="ml"><i class="l-not"></i> ≥ 30%</span> <span class="ml"><i class="l-ext"></i> ≥ 98% ou zerada</span> <span class="ml"><i class="l-new"></i> verba nova</span> <span class="ml"><span class="l-trend"></span> reta de tendência</span> <span class="ml"><i class="l-out"></i> fora da tendência</span></div>
      <div class="table-wrap hol-scroll" id="holScroll" style="max-height:460px">
        <table class="data-table hol-table"><thead>${head}</thead><tbody id="holBody">${body}</tbody></table>
      </div>`;
    _bindHolInner();
    const table = document.querySelector('#holContent .hol-table');
    [...table.tBodies[0].rows].forEach((tr, i) => sum.view.bind(tr, sum.page[i]));
    sum.view.finish(table);
  }
  function _holValSort(l, key) {
    if (key === "calc") return _holLineChecked(l) ? 1 : 0;
    if (key === "cod")  return String(l.codigo);
    if (key === "desc") return String(l.descricao);
    if (key === "clas") return _clasRank(l.clas);
    if (key && key.indexOf("m:") === 0) return l.per[key.slice(2)] || 0;
    if (key === "trend") {
      const seq = S.mesesDetectados.map(m => l.per[m] || 0);
      const f = seq[0], last = seq[seq.length - 1];
      return Math.abs(f) > 0.005 ? (last - f) / Math.abs(f) : (Math.abs(last) > 0.005 ? 1 : 0);
    }
    return "";
  }
  function _bindHolInner() {
    const busca = document.getElementById("holBusca");
    if (busca) busca.addEventListener("input", e => { HOL.q = e.target.value; _holReRender(); });
    const onlyc = document.getElementById("holOnlyChk");
    if (onlyc) onlyc.addEventListener("change", e => { HOL.onlyChk = e.target.checked; _holReRender(); });
    document.querySelectorAll("#holContent .hol-chip").forEach(ch => ch.addEventListener("click", () => {
      const k = ch.dataset.clas; HOL.clas[k] = !HOL.clas[k]; _holReRender();
    }));
    const body = document.getElementById("holBody");
    if (body) body.addEventListener("change", e => {
      const chk = e.target.closest(".hol-check"); if (!chk) return;
      _confSetLinha(_confKeyMN(HOL.mat, HOL.nome), chk.dataset.lk, chk.checked);
      _holReRender();
      if (S.funcSelecionado && S.funcSelecionado.Matrícula === HOL.mat && S.funcSelecionado.Nome === HOL.nome) {
        _renderConfBox(S.funcSelecionado); _renderizarDrillVerbas();
      }
      _renderConferencia();
    });
    _bindSort("#tab-holerite .hol-table", HOL.sort, _holReRender);
  }
  function _holReRender() {
    const sc = document.getElementById("holScroll"); const top = sc ? sc.scrollTop : 0;
    const fb = document.activeElement && document.activeElement.id === "holBusca";
    _renderHolerite();
    const sc2 = document.getElementById("holScroll"); if (sc2) sc2.scrollTop = top;
    if (fb) { const b = document.getElementById("holBusca"); if (b) { b.focus(); b.selectionStart = b.value.length; } }
  }
  function _holNav(delta) {
    const sel = document.getElementById("holFunc"); if (!sel || !sel.options.length) return;
    let i = sel.selectedIndex + delta;
    i = Math.max(0, Math.min(sel.options.length - 1, i));
    sel.selectedIndex = i;
    const [emp, m, n] = JSON.parse(sel.value); HOL.empresa = emp; HOL.mat = m; HOL.nome = n; HOL.q = ""; HOL.onlyChk = false;
    _renderHolerite();
  }
  function _verHolerite(mat, nome, empresa) {
    HOL.empresa = empresa || S.funcSelecionado?.Empresa || (S.liquido || []).find(r => r.Matrícula === mat && r.Nome === nome)?.Empresa;
    HOL.mat = mat; HOL.nome = nome; HOL.q = ""; HOL.onlyChk = false; HOL.buscaColab = "";
    const hb = document.getElementById("holBuscaColab"); if (hb) hb.value = "";
    _switchTab("holerite");
    _renderHolerite();
  }

  function _carregarEmpresas() {
    const emps = [...new Set(S.df.map(r => (r["Empresa"] || "").trim()).filter(Boolean))].sort();
    S.empresasDisp = emps; S.empresasSel = [...emps];
    const cont = document.getElementById("filtroEmpresa");
    if (emps.length <= 1) { cont.innerHTML = `<div class="hint-text" style="margin-top:4px">Empresa: ${emps[0] || "-"}</div>`; return; }
    cont.innerHTML = `<label class="sidebar-label">Empresa(s)</label>
      <div class="sel-row">
        <button class="btn-xs" onclick="APP._selTodosEmpresa(true)">Todas</button>
        <button class="btn-xs" onclick="APP._selTodosEmpresa(false)">Nenhuma</button>
        <span style="font-size:9px;color:#999;margin-left:2px" id="empCount">${emps.length}/${emps.length}</span>
      </div>
      <div class="check-frame" id="empFrame"></div>`;
    const frame = document.getElementById("empFrame");
    emps.forEach(emp => {
      const lbl = document.createElement("label");
      lbl.innerHTML = `<input type="checkbox" checked value="${emp}" data-type="empresa"> ${emp}`;
      lbl.querySelector("input").addEventListener("change", () => {
        S.empresasSel = [...document.querySelectorAll("[data-type=empresa]:checked")].map(cb => cb.value);
        const n = S.empresasSel.length;
        const el = document.getElementById("empCount"); if (el) el.textContent = `${n}/${emps.length}`;
      });
      frame.appendChild(lbl);
    });
  }

  function _selTodosEmpresa(val) {
    document.querySelectorAll("[data-type=empresa]").forEach(cb => cb.checked = val);
    S.empresasSel = val ? [...S.empresasDisp] : [];
    const el = document.getElementById("empCount"); if (el) el.textContent = `${S.empresasSel.length}/${S.empresasDisp.length}`;
  }

  function _carregarProcessos() {
    const procs = [...new Set(S.df.map(r => (r["Processo"] || "").trim()).filter(Boolean))].sort();
    S.processosDisp = procs; S.processosSel = [...procs];
    const cont = document.getElementById("filtroProcesso");
    if (procs.length <= 1) { cont.innerHTML = procs.length ? `<div class="hint-text">Processo: ${procs[0]}</div>` : ""; return; }
    cont.innerHTML = `<label class="sidebar-label">Processo (tipo de folha)</label>
      <div class="sel-row">
        <button class="btn-xs" onclick="APP._selTodosProcesso(true)">Todos</button>
        <button class="btn-xs" onclick="APP._selTodosProcesso(false)">Nenhum</button>
        <span style="font-size:9px;color:#999;margin-left:2px" id="procCount">${procs.length}/${procs.length}</span>
      </div>
      <div class="check-frame" id="procFrame"></div>`;
    const frame = document.getElementById("procFrame");
    procs.forEach(p => {
      const lbl = document.createElement("label");
      lbl.innerHTML = `<input type="checkbox" checked value="${p}" data-type="processo"> ${p}`;
      lbl.querySelector("input").addEventListener("change", () => {
        S.processosSel = [...document.querySelectorAll("[data-type=processo]:checked")].map(cb => cb.value);
        const el = document.getElementById("procCount"); if (el) el.textContent = `${S.processosSel.length}/${procs.length}`;
      });
      frame.appendChild(lbl);
    });
  }

  function _selTodosProcesso(val) {
    document.querySelectorAll("[data-type=processo]").forEach(cb => cb.checked = val);
    S.processosSel = val ? [...S.processosDisp] : [];
    const el = document.getElementById("procCount"); if (el) el.textContent = `${S.processosSel.length}/${S.processosDisp.length}`;
  }

  function _popularComboMesAlvo() {
    const sel = document.getElementById("comboMesAlvo"); sel.disabled = false; sel.innerHTML = "";
    S.mesesDetectados.forEach(m => { const o = document.createElement("option"); o.value = m; o.textContent = m; sel.appendChild(o); });
    sel.value = S.mesesDetectados[S.mesesDetectados.length - 1];
    S.mesAlvo = sel.value;
    sel.addEventListener("change", () => { S.mesAlvo = sel.value; _renderBaselineCheckboxes(); });
  }

  function _renderBaselineCheckboxes() {
    const frame = document.getElementById("baselineFrame"); frame.innerHTML = "";
    S.checksBaseline = {};
    S.mesesDetectados.filter(m => m !== S.mesAlvo).forEach(m => {
      const lbl = document.createElement("label");
      const cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = true; cb.value = m;
      cb.addEventListener("change", _syncBaseline);
      lbl.appendChild(cb); lbl.appendChild(document.createTextNode(" " + m));
      frame.appendChild(lbl); S.checksBaseline[m] = cb;
    });
    _syncBaseline();
  }

  function _syncBaseline() {
    S.mesesBaseline = S.mesesDetectados.filter(m => S.checksBaseline[m] && S.checksBaseline[m].checked);
  }

  function _onMetodoChange() {
    S.configNome = document.getElementById("comboMetodo").value;
    document.getElementById("lblMetodoDesc").textContent = PARAMETROS_MERCADO[S.configNome].descricao;
  }

  function _onMatChange(e) {
    S.matPct = parseFloat(e.target.value);
    document.getElementById("lblMatValor").textContent = `${S.matPct.toFixed(1)}%`;
  }

  // ---- Reprocessar ----
  function acaoReprocessar() {
    if (!S.df) return;

    // Modo mês único: se só há um mês no arquivo, segue sem baseline (sem comparativos).
    // Se há vários meses mas o usuário desmarcou todos, aí sim é erro do usuário.
    const modoMesUnico = !S.mesesBaseline.length;
    if (modoMesUnico && S.mesesDetectados.length > 1) {
      alert("Selecione ao menos um mês de baseline (ou use um arquivo com um único mês para o modo sem histórico).");
      return;
    }
    S.modoMesUnico = modoMesUnico;

    let dfF = S.df.filter(r => S.empresasSel.includes((r["Empresa"] || "").trim()));
    if (S.processosDisp.length > 1) dfF = dfF.filter(r => S.processosSel.includes((r["Processo"] || "").trim()));


    if (S.verbasExcluidas.size > 0) {
      const exclNorm = new Set([...S.verbasExcluidas].map(normCod));
      exclNorm.delete("20"); // verba 0020 sempre protegida
      dfF = dfF.filter(r => !exclNorm.has(normCod(r["Código"])));
    }
    S.dfFiltrado = dfF;
    S.sumLiquids = new Map();
    for (const row of dfF) {
      if (_isLiqRow(row) || !['PGTO','DESC'].includes(row['Clas.'])) continue;
      const n = _sumPayroll(row)['Folha (proventos − descontos)'] || 0n;
      S.sumLiquids.set(funcKey(row), (S.sumLiquids.get(funcKey(row)) || 0n) + n);
    }

    // ---- Constrói mapa de anonimização (mat → {mat anon, nome anon}) ----
    const matsOrdenadas = [...new Set(S.df.map(getMatricula))].sort();
    S.mapaAnonimo = {};
    matsOrdenadas.forEach((mat, i) => {
      const n = String(i + 1).padStart(3, "0");
      S.mapaAnonimo[mat] = { mat: `MAT-${n}`, nome: `Colaborador ${n}` };
    });
    window.AF_ANON = S.anonimizar;
    window.AF_MAP  = S.mapaAnonimo;

    const cfg = Object.assign({}, PARAMETROS_MERCADO[S.configNome]);
    cfg.materialidade_pct = S.matPct / 100;

    const desligados = detectarDesligados(dfF, S.mesAlvo);
    S.resumo     = resumoMacro(dfF, S.mesesDetectados);
    const { result: liq, metadata } = liquidoPorFuncionario(dfF, S.mesesDetectados, S.mesAlvo, S.mesesBaseline, desligados, cfg, modoMesUnico);
    S.liquido    = liq;
    S.metadata   = metadata;
    S.descontosRevisar = descontosParaRevisar(dfF, S.mesesDetectados, S.mesAlvo, S.mesesBaseline, liq, PARAMS_REVISAO);
    S.salarioV20 = salarioVerba20PorMes(dfF, S.mesesDetectados);
    S.zeradosPGTO = verbasZeradas(dfF, "PGTO", S.mesAlvo, S.mesesBaseline);
    S.zeradosDESC = verbasZeradas(dfF, "DESC", S.mesAlvo, S.mesesBaseline);
    S.liquidoAnalise = S.excluirDesligados ? liq.filter(r => r.STATUS !== "DESLIGADO") : liq;

    const nFunc = new Set(dfF.map(getMatricula)).size;
    document.getElementById("lblStats").textContent =
      `${dfF.length.toLocaleString("pt-BR")} linhas | ${nFunc} func.\n` +
      `${S.empresasSel.length} empresa(s) | ${S.mesesBaseline.length} baseline\n` +
      `Folha ${S.mesAlvo}: ${fmtBRL(metadata.folhaTotalAlvo)}\n` +
      `Materialidade: ${fmtBRL(metadata.materialidade)}`;
    const lvEl = document.getElementById("lblVerbasStatus");
    if (S.verbasExcluidas.size > 0) { lvEl.textContent = `${S.verbasExcluidas.size} verba(s) excluída(s)`; lvEl.style.color = IG_VINHO; }
    else { lvEl.textContent = "(todas as verbas incluídas)"; lvEl.style.color = ""; }

    _renderVisaoGeral();
    _renderHeadcount();
    // Outliers e Análise de Verbas: lazy (canvas em aba inativa = dimensão zero)
    const tabAtiva = document.querySelector(".tab-btn.active")?.dataset.tab;
    if (tabAtiva === "outliers")      _renderOutliers(); else { S.tabsDirty["outliers"] = true; }
    if (tabAtiva === "analiseVerbas") _renderAnaliseVerbas(); else { S.tabsDirty["analiseVerbas"] = true; }
    _renderVerbasZeradas();
    _popularCombosTabela();
    // Bind sort por cabeçalho (só uma vez, após tabela existir)
    _bindSortHeaders();
    _renderizarTabela();
    _renderConferencia();
    _renderHolerite();
    if (S.funcSelecionado) {
      S.funcSelecionado = S.liquidoAnalise.find(r => funcKey(r) === funcKey(S.funcSelecionado)) || null;
    }
    if (S.funcSelecionado) _renderizarDrillVerbas();
    else {
      document.getElementById('drillBody').replaceChildren();
      const empty = _sumView('verbas-detalhe', [], {values: _sumPayroll, render: _renderizarDrillVerbas, key: row => RowSelection.sourceId(row)});
      empty.view.finish(document.getElementById('tabelaDrill'));
      document.getElementById('drillInfo').textContent = '(duplo clique em um funcionário na tabela acima)';
    }
    let av = S.df.filter(r => S.empresasSel.includes(getEmpresa(r)));
    if (S.processosDisp.length > 1) av = av.filter(r => S.processosSel.includes((r['Processo'] || '').trim()));
    S.dfAvBase = av;
    if (S.colabDfAtual) _popularTabelaColab(S.codigosColabAtual);

  }

  // ---- Helpers de anonimização ----
  function maskNome(nome, mat) {
    if (!S.anonimizar || !S.mapaAnonimo) return nome;
    return S.mapaAnonimo[mat]?.nome || nome;
  }
  function maskMat(mat) {
    if (!S.anonimizar || !S.mapaAnonimo) return mat;
    return S.mapaAnonimo[mat]?.mat || mat;
  }

  // ---- Abas de renderização ----
  function _renderVisaoGeral() {
    const la = S.resumo.find(r => r.Mes === S.mesAlvo);
    const lb = S.resumo.filter(r => S.mesesBaseline.includes(r.Mes));
    const semHist = S.modoMesUnico || !lb.length;
    const _med = (fn) => lb.length ? mean(lb.map(fn)) : 0;
    const mPGTO = _med(r=>r.PGTO), mDESC = _med(r=>r.DESC), mLiq = _med(r=>r.Liquido);
    const vP = mPGTO ? ((la.PGTO/mPGTO)-1)*100 : 0;
    const vD = mDESC ? ((la.DESC/mDESC)-1)*100 : 0;
    const vL = mLiq  ? ((la.Liquido/mLiq)-1)*100 : 0;
    const diff9950 = la.Liquido - la.Verba9950;
    const pct9950  = la.Verba9950 ? (diff9950/la.Verba9950)*100 : 0;
    const sub = (v) => semHist ? "sem histórico" : `${fmtPct(v)} vs baseline`;
    _card("cardPGTO", `PGTO ${S.mesAlvo}`, fmtBRL(la.PGTO),   sub(vP), (!semHist && vP < -10) ? "alerta" : "");
    _card("cardDESC", `DESC ${S.mesAlvo}`, fmtBRL(la.DESC),   sub(vD), (!semHist && vD < -10) ? "alerta" : "");
    _card("cardLIQ",  `Líquido ${S.mesAlvo}`, fmtBRL(la.Liquido), sub(vL), (!semHist && vL < -10) ? "alerta" : "");
    _card("card9950", "Verba 9950", fmtBRL(la.Verba9950), `Diff: ${fmtBRL(diff9950)} (${fmtPct(pct9950)})`, Math.abs(pct9950) > 15 ? "alerta" : "");
    const nDesl = S.liquido.filter(r => r.STATUS === "DESLIGADO").length;
    const avisos = [];
    if (semHist) avisos.push("Modo mês único: sem histórico disponível — comparativos desativados.");
    if (nDesl) avisos.push(`Desligados detectados: ${nDesl} | ${S.excluirDesligados ? "excluídos da análise de líquido" : "incluídos"}.`);
    document.getElementById("avisoDesligados").textContent = avisos.join(" ");
    renderEvolucaoMensal("chartEvolucao", S.resumo, S.mesAlvo, S.mesesBaseline);
    renderConciliacao("chartConciliacao", la, S.mesAlvo);
    renderDistribuicaoStatus("chartStatus", S.liquidoAnalise);
    _renderPainelDescontos();
  }

  // ---- Painel "Descontos para revisar" (Visão Geral) ----
  function _filtraDescontos() {
    const q = (S.vgBusca || "").trim().toLowerCase();
    let lista = S.descontosRevisar || [];
    if (q) lista = lista.filter(e =>
      String(e.Matrícula).toLowerCase().includes(q) ||
      String(e.Nome).toLowerCase().includes(q) ||
      String(e.Código).toLowerCase().includes(q) ||
      String(e.Descrição).toLowerCase().includes(q));
    return lista;
  }

  function _renderPainelDescontos() {
    const wrap = document.getElementById("painelDescontos");
    if (!wrap) return;
    const todos = S.descontosRevisar || [];
    let lista = _filtraDescontos();
    // Ordenação por clique no cabeçalho (asc/desc). Sem clique, mantém a ordem
    // padrão (prioridade -> comprometimento -> valor) definida em descontosParaRevisar.
    if (S.descSort.key) lista = lista.slice().sort(_genCmp(_descValSort, S.descSort.key, S.descSort.dir));
    // KPIs (sempre sobre a população completa — o filtro é só de exibição)
    const obrig = todos.filter(e => e.PRIORIDADE === "Revisão obrigatória");
    const totalDesc = todos.reduce((s, e) => s + Math.abs(e.VALOR), 0);
    const nColab = new Set(todos.map(e => e.funcKey)).size;
    const kpis = document.getElementById("descKpis");
    if (kpis) kpis.innerHTML =
      _kpi("Colaboradores afetados", fmtInt(nColab)) +
      _kpi("Descontos sinalizados", fmtInt(todos.length)) +
      _kpi("Revisão obrigatória", fmtInt(obrig.length), obrig.length ? "alerta" : "") +
      _kpi("Total sinalizado", fmtBRL(totalDesc));

    // Gráfico de dispersão (valor x comprometimento) — casos similares
    if (typeof renderScatterDescontos === "function") renderScatterDescontos("chartDescontos", todos);

    // Tabela sem cores (Empresa | Matrícula | Nome | Verba | Valor | Referência | Impacto % | Prioridade | Motivo)
    const body = document.getElementById("descBody");
    if (!body) return;
    const sum = _sumView('descontos', lista, {
      key: e => JSON.stringify([e.funcKey, normCod(e.Código), S.mesAlvo]),
      values: e => RowSelection.payroll(e.VALOR, 'DESC', e.Código, e.Descrição), emptyGroup: 'Folha (proventos − descontos)',
      render: _renderPainelDescontos, investigate: e => _verHolerite(e.Matrícula, e.Nome, e.Empresa)
    });
    if (!lista.length) {
      body.innerHTML = `<tr><td colspan="9" class="hint-text" style="text-align:center;padding:14px">${
        todos.length ? "Nenhum desconto para o filtro atual." : "Nenhum desconto sinalizado para revisão neste mês."}</td></tr>`;
    } else {
      body.innerHTML = sum.page.map(e => {
        const nome = maskNome(e.Nome, e.Matrícula), mat = maskMat(e.Matrícula);
        const comp = e.COMPROMETIMENTO != null ? `${e.COMPROMETIMENTO.toFixed(1)}%` : "-";
        const prioCls = e.PRIORIDADE === "Revisão obrigatória" ? "prio-obrig" : "prio-alerta";
        return `<tr data-funckey="${escAttr(e.funcKey)}" class="desc-row" title="Duplo clique abre o holerite">
          <td>${escHtml(e.Empresa)}</td>
          <td class="center">${escHtml(mat)}</td>
          <td>${escHtml(nome)}</td>
          <td>${escHtml(String(e.Código))} — ${escHtml(e.Descrição || "")}</td>
          <td class="num">${fmtBRL(e.VALOR)}</td>
          <td class="num">${e.REFERENCIA ? fmtBRL(e.REFERENCIA) : "—"}</td>
          <td class="num">${comp}</td>
          <td><span class="prio-badge ${prioCls}">${e.PRIORIDADE}</span></td>
          <td style="font-size:11px">${escHtml(e.MOTIVO)}</td>
        </tr>`;
      }).join("");
    }
    if (sum.page.length) [...body.rows].forEach((tr, i) => sum.view.bind(tr, sum.page[i]));
    sum.view.finish(document.getElementById('tabelaDescontos'));
    document.getElementById("descCount").textContent =
      `${lista.length} de ${todos.length} desconto(s)` + ` (${sum.page.length} exibidos)`;
    _bindSort("#tabelaDescontos", S.descSort, _renderPainelDescontos);
  }

  // Acessor de valor para ordenação da tabela de descontos
  function _descValSort(e, key) {
    switch (key) {
      case "Matrícula": { const n = parseFloat(String(e.Matrícula).replace(/\D/g, "")); return isNaN(n) ? String(e.Matrícula) : n; }
      case "Código":    { const n = parseFloat(String(e.Código).replace(/\D/g, "")); return isNaN(n) ? String(e.Código) : n; }
      case "VALOR":          return Math.abs(e.VALOR || 0);
      case "REFERENCIA":     return Math.abs(e.REFERENCIA || 0);
      case "COMPROMETIMENTO": return e.COMPROMETIMENTO == null ? -1 : e.COMPROMETIMENTO;
      case "PRIORIDADE":     return e.PRIORIDADE === "Revisão obrigatória" ? 0 : 1;
      default:               return e[key] != null ? e[key] : "";
    }
  }

  function _kpi(label, valor, cls = "") {
    return `<div class="desc-kpi ${cls}"><div class="desc-kpi-label">${label}</div><div class="desc-kpi-valor">${valor}</div></div>`;
  }

  function _renderHeadcount() {
    const la = S.salarioV20.find(r => r.Mes === S.mesAlvo);
    const lb = S.salarioV20.filter(r => S.mesesBaseline.includes(r.Mes));
    const hcMed  = lb.length ? mean(lb.map(r=>r.HC)) : 0;
    const salMed = lb.length ? mean(lb.map(r=>r.Salario_Medio)) : 0;
    const vH = hcMed  ? ((la.HC/hcMed)-1)*100 : 0;
    const vS = salMed ? ((la.Salario_Medio/salMed)-1)*100 : 0;
    _card("cardHC",    `HC verba 0020 em ${S.mesAlvo}`, fmtInt(la.HC), `${fmtPct(vH)} vs baseline`, Math.abs(vH)>8?"alerta":"");
    _card("cardSal",   `Salário médio em ${S.mesAlvo}`, fmtBRL(la.Salario_Medio), `${fmtPct(vS)} vs baseline`, Math.abs(vS)>8?"alerta":"");
    _card("cardTotal20", `Total verba 0020 — ${S.mesAlvo}`, fmtBRL(la.Salario_Total), "", "");
    renderHeadcount("chartHC", S.salarioV20, S.mesAlvo);
  }

  function _renderOutliers() {
    renderDispersao("chartDispersao", S.liquidoAnalise, S.mesAlvo, (point) => {
      _navigateToFuncDrill(point.mat, point.nome);
    });
    renderTopOutliers("chartTopOutliers", S.liquidoAnalise, S.mesAlvo);
  }

  function _renderVerbasZeradas() {
    renderVerbasZeradas("chartZeradosPGTO", S.zeradosPGTO, `PGTO zeradas em ${S.mesAlvo}`, IG_AZUL);
    renderVerbasZeradas("chartZeradosDESC", S.zeradosDESC, `DESC zeradas em ${S.mesAlvo}`, IG_VINHO);
  }

  // ---- Navegação: painel de descontos → Holerite ----
  function _abrirHolerite(mat, nome) {
    _switchTab("holerite");
    HOL.mat = mat; HOL.nome = nome; HOL.buscaColab = String(mat);
    HOL.q = ""; HOL.onlyChk = false;
    const hb = document.getElementById("holBuscaColab"); if (hb) hb.value = String(mat);
    _renderHolerite();
  }

  // ---- Navegação: scatter → Tabela & Drill-down ----
  function _navigateToFuncDrill(mat, nome, empresa) {
    _switchTab("tabela");
    const r = (S.liquidoAnalise || []).find(x => x.Matrícula === mat && x.Nome === nome && (!empresa || x.Empresa === empresa));
    if (!r) return;
    S.funcSelecionado = r;
    _renderDrill(r);
    document.getElementById('drillInfo')?.scrollIntoView({behavior:'smooth', block:'nearest'});

  }

  // ---- Drill colaborador na aba Análise de Verbas ----
  function _renderColabDrill(r) {
    const infoEl = document.getElementById("colabDrillInfo");
    const cods = S.codigosColabAtual;

    // Constrói série com os valores da(s) verba(s) selecionada(s) para este colaborador
    if (!cods.length || !S.dfAvBase) {
      if (infoEl) infoEl.textContent = "Nenhuma verba selecionada.";
      return;
    }
    const alvo = new Set(cods.map(normCod));
    const sub  = S.dfAvBase.filter(row =>
      alvo.has(normCod(row["Código"])) &&
      funcKey(row) === funcKey(r)
    );

    const serie = { STATUS: r.STATUS_LIQ };
    for (const m of S.mesesDetectados) {
      const col = `${m} - Valor`;
      serie[m] = sub.reduce((sum, row) => sum + (row[col] || 0), 0);
    }
    const baseVals = S.mesesBaseline.map(m => serie[m]).filter(v => v !== 0);
    serie.MEDIA_BASELINE = baseVals.length ? mean(baseVals) : null;

    // Monta info-bar com botão de navegação inline
    if (infoEl) {
      infoEl.innerHTML = "";
      const dNome = maskNome(r.Nome, r.Matrícula);
      const dMat  = maskMat(r.Matrícula);
      const txt = document.createTextNode(
        `Mat. ${dMat} — ${dNome}  |  Status: ${r.STATUS_LIQ}  |  ` +
        `Verba ${S.mesAlvo}: ${fmtBRL(r.VALOR_ALVO)}  |  Líquido: ${fmtBRL(r.LIQUIDO_ALVO)} (${fmtPct(r.VAR_PCT)})`
      );
      const btn = document.createElement("button");
      btn.textContent = "→ Abrir drill-down completo";
      btn.className = "btn-sm btn-primary";
      btn.style.cssText = "margin-left:14px;padding:2px 10px;font-size:11px;display:inline-block;width:auto";
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        _navigateToFuncDrill(r.Matrícula, r.Nome, r.Empresa);
      });
      infoEl.appendChild(txt);
      infoEl.appendChild(btn);
    }

    // Label: nome das verbas selecionadas
    const labelVerba = cods.length === 1
      ? `Verba ${cods[0]}`
      : `${cods.length} verbas: ${cods.slice(0, 3).join(", ")}${cods.length > 3 ? "..." : ""}`;

    renderDrillFuncionario("chartColabDrill", serie, S.mesesDetectados, S.mesAlvo, serie.MEDIA_BASELINE, labelVerba);
  }
  function _renderAnaliseVerbas() {
    let dfAV = S.df.filter(r => S.empresasSel.includes((r["Empresa"] || "").trim()));
    if (S.processosDisp.length > 1) dfAV = dfAV.filter(r => S.processosSel.includes((r["Processo"] || "").trim()));
    S.dfAvBase = dfAV;
    S.topVerbasDF = impactoPorVerba(dfAV, S.mesesDetectados, S.mesAlvo, S.mesesBaseline, 20);
    renderTopVerbas("chartTopVerbas", S.topVerbasDF, S.mesAlvo, S.verbasExcluidas);

    // Agrega por Código+Descrição+Clas para listbox
    const valCols = S.mesesDetectados.map(m => `${m} - Valor`);
    const agg = {};
    for (const r of dfAV) {
      const key = `${r["Código"]}|||${r["Descrição"]}|||${r["Clas."]}`;
      if (!agg[key]) agg[key] = { Código: r["Código"], Descrição: r["Descrição"], "Clas.": r["Clas."], totalAbs: 0 };
      for (const c of valCols) agg[key].totalAbs += Math.abs(r[c] || 0);
    }
    const ordemClas = { PGTO: 0, DESC: 1, OUTRO: 2 };
    S.verbasAvCompleta = Object.values(agg).filter(d => d.totalAbs > 0).sort((a, b) => {
      const oa = ordemClas[a["Clas."]] ?? 99, ob = ordemClas[b["Clas."]] ?? 99;
      return oa !== ob ? oa - ob : b.totalAbs - a.totalAbs;
    });
    _filtrarListboxVerbas();
    const lb = document.getElementById("listboxVerbas");
    if (lb.options.length > 0 && lb.selectedIndex < 0) lb.selectedIndex = 0;
    _renderDrillVerba();
  }

  function _filtrarListboxVerbas() {
    if (!S.verbasAvCompleta) return;
    const lb = document.getElementById("listboxVerbas");
    const busca = (document.getElementById("entryBuscaVerba").value || "").toUpperCase().trim();
    const selCods = new Set([...lb.options].filter(o => o.selected).map(o => o.value));
    lb.innerHTML = "";
    let df = S.verbasAvCompleta;
    if (busca) df = df.filter(r =>
      String(r.Código).toUpperCase().includes(busca) ||
      String(r.Descrição).toUpperCase().includes(busca) ||
      String(r["Clas."]).toUpperCase().includes(busca)
    );
    for (const r of df) {
      const o = document.createElement("option");
      o.value = String(r.Código);
      o.textContent = `[${String(r["Clas."]).padEnd(5)}] ${String(r.Código).padStart(7)} — ${String(r.Descrição).substring(0,55)}`;
      o.selected = selCods.has(String(r.Código));
      lb.appendChild(o);
    }
  }

  function _extrairCodsListbox() {
    return [...document.getElementById("listboxVerbas").options].filter(o => o.selected).map(o => o.value);
  }

  function selecionarClasseVerbas(classe) {
    const lb = document.getElementById("listboxVerbas");
    [...lb.options].forEach(o => { o.selected = o.textContent.startsWith(`[${classe}`) || o.textContent.startsWith(`[${classe.padEnd(5)}]`); });
    _renderDrillVerba();
  }

  function limparSelecaoVerbas() {
    [...document.getElementById("listboxVerbas").options].forEach(o => o.selected = false);
    _renderDrillVerba();
  }

  function _renderDrillVerba() {
    if (!S.dfAvBase) return;
    const cods = _extrairCodsListbox();
    const infoEl = document.getElementById("verbasInfo");
    if (!cods.length) { infoEl.textContent = "(nenhuma selecionada)"; _limparCardsAV(); _popularTabelaColab([]); return; }
    infoEl.textContent = `${cods.length} verba(s):\n${cods.slice(0,8).join("\n")}${cods.length>8 ? `\n...+${cods.length-8}` : ""}`;
    const dados = hcETotalPorVerba(S.dfAvBase, S.mesesDetectados, cods);
    const la = dados.find(d => d.Mes === S.mesAlvo);
    const lb = dados.filter(d => S.mesesBaseline.includes(d.Mes));
    const hcMed  = lb.length ? mean(lb.map(d=>d.HC)) : 0;
    const medMed = lb.length ? mean(lb.map(d=>d.Media)) : 0;
    const totMed = lb.length ? mean(lb.map(d=>d.Total)) : 0;
    const vH = hcMed  ? ((la.HC/hcMed)-1)*100 : 0;
    const vM = medMed ? ((la.Media/medMed)-1)*100 : 0;
    const vT = totMed ? ((la.Total/totMed)-1)*100 : 0;
    _card("cardAVHC",    `HC ${S.mesAlvo}`, fmtInt(la.HC), `${fmtPct(vH)} vs baseline`, Math.abs(vH)>15?"alerta":"");
    _card("cardAVTotal", `Total ${S.mesAlvo}`, fmtBRL(la.Total), `${fmtPct(vT)} vs baseline`, "");
    _card("cardAVMedia", "Valor médio / matrícula", fmtBRL(la.Media), `${fmtPct(vM)} vs baseline`, "");
    const exclCount = cods.filter(c => S.verbasExcluidas.has(c)).length;
    _card("cardAVStatus", "Status no cálculo",
      exclCount === 0 ? "Todas incluídas" : exclCount === cods.length ? "Todas excluídas" : `${exclCount}/${cods.length} excluídas`,
      `${cods.length} verba(s)`, exclCount > 0 ? "atencao" : "");
    const titulo = cods.length === 1 ? cods[0] : `${cods.length} verbas: ${cods.slice(0,4).join(", ")}`;
    renderAVHC("chartAVHC", dados, S.mesAlvo, S.mesesBaseline, titulo);
    _popularTabelaColab(cods);
  }

  function _limparCardsAV() {
    ["cardAVHC","cardAVTotal","cardAVMedia","cardAVStatus"].forEach(id => {
      const el = document.getElementById(id); if (el) el.innerHTML = "";
    });
  }

  // ---- Tabela colaboradores ----
  function _popularTabelaColab(cods) {
    S.codigosColabAtual = cods;
    if (!cods.length) { S.colabDfAtual = []; _renderizarTabelaColab(); return; }
    const alvo = new Set(cods.map(normCod));
    const sub = S.dfAvBase.filter(r => alvo.has(normCod(r["Código"])));
    const funcMap = {};
    for (const r of sub) {
      const key = funcKey(r);
      if (!funcMap[key]) funcMap[key] = { empresa: getEmpresa(r), mat: getMatricula(r), nome: getNome(r), vAlvo: 0, totalPorMes: {}, source: [] };
      funcMap[key].source.push(r);
      funcMap[key].vAlvo += (r[`${S.mesAlvo} - Valor`] || 0);
      // Agrega por mês (soma todas as verbas selecionadas do funcionário no mês)
      for (const m of S.mesesBaseline) {
        funcMap[key].totalPorMes[m] = (funcMap[key].totalPorMes[m] || 0) + (r[`${m} - Valor`] || 0);
      }
    }
    const colab = [];
    for (const [key, d] of Object.entries(funcMap)) {
      const baseVals = Object.values(d.totalPorMes);
      const nonZeroBase = baseVals.filter(v => v !== 0);
      const mediaVerba = nonZeroBase.length ? mean(nonZeroBase) : null;
      const varAbs = mediaVerba != null ? d.vAlvo - mediaVerba : null;
      const varPct = mediaVerba ? ((d.vAlvo / mediaVerba) - 1) * 100 : null;
      const liqRow = S.liquido ? S.liquido.find(r => r.Matrícula === d.mat && r.Nome === d.nome && r.Empresa === d.empresa) : null;
      colab.push({ Empresa: d.empresa, source: d.source, Matrícula: d.mat, Nome: d.nome, STATUS_LIQ: liqRow?.STATUS || "-",
        VALOR_ALVO: d.vAlvo, MEDIA_VERBA: mediaVerba, VAR_ABS: varAbs, VAR_PCT: varPct,
        LIQUIDO_ALVO: liqRow?.LIQUIDO_ALVO, AUDITORIA: liqRow?.AUDITORIA || "" });
    }
    S.colabDfAtual = colab;
    const statuses = [...new Set(colab.map(r => r.STATUS_LIQ))].sort();
    const comboColabStatus = document.getElementById("comboColabStatus");
    const valorAtual = comboColabStatus.value;
    comboColabStatus.innerHTML = "<option>(todos)</option>";
    statuses.forEach(s => { const o = document.createElement("option"); o.value = s; o.textContent = s; comboColabStatus.appendChild(o); });
    if (statuses.includes(valorAtual)) comboColabStatus.value = valorAtual;
    _renderizarTabelaColab();
  }

  function _colabValSort(r, key) {
    switch (key) {
      case "mat":    return String(r.Matrícula);
      case "nome":   return String(r.Nome);
      case "status": return String(r.STATUS_LIQ);
      case "valor":  return r.VALOR_ALVO || 0;
      case "media":  return r.MEDIA_VERBA || 0;
      case "var":    return Math.abs(r.VAR_ABS || 0);
      case "varpct": return r.VAR_PCT == null ? -Infinity : r.VAR_PCT;
      case "liq":    return r.LIQUIDO_ALVO || 0;
      case "aud":    return String(r.AUDITORIA || "");
      case "conf":   { const c = S.conf[_confKeyMN(r.Matrícula, r.Nome)]; return (c && c.ok) ? 1 : 0; }
      default:       return "";
    }
  }
  function _renderizarTabelaColab() {
    if (!S.colabDfAtual) return;
    let df = [...S.colabDfAtual];
    const stFilt = document.getElementById("comboColabStatus").value;
    if (stFilt && stFilt !== "(todos)") df = df.filter(r => r.STATUS_LIQ === stFilt);
    const busca = (document.getElementById("entryColabBusca").value || "").toUpperCase().trim();
    if (busca) df = df.filter(r => r.Nome.toUpperCase().includes(busca) || String(r.Matrícula).includes(busca));
    if (S.colabSort.key) df.sort(_genCmp(_colabValSort, S.colabSort.key, S.colabSort.dir));
    else df.sort((a, b) => Math.abs(b.VAR_ABS || 0) - Math.abs(a.VAR_ABS || 0));
    const tbody = document.getElementById("colabBody"); tbody.innerHTML = "";
    const sum = _sumView('colaboradores-verbas', df, {
      key: r => JSON.stringify([funcKey(r), S.mesAlvo, [...S.codigosColabAtual].sort()]),
      values: r => RowSelection.combine(r.source.map(_sumPayroll)), emptyGroup: 'Folha (proventos − descontos)', label: 'Composição das verbas · ' + S.mesAlvo,
      render: _renderizarTabelaColab, investigate: r => {
        _renderColabDrill(r); document.getElementById('colabDrillInfo')?.scrollIntoView({behavior:'smooth', block:'nearest'});
      }
    });
    sum.page.forEach(r => {
      const tr = document.createElement("tr");
      if (STATUS_CRITICOS.has(r.STATUS_LIQ)) tr.className = "critico";
      const dNome = maskNome(r.Nome, r.Matrícula);
      const dMat  = maskMat(r.Matrícula);
      tr.innerHTML = `<td class="center">${dMat}</td><td>${String(dNome).substring(0,55)}</td>
        <td class="center"><span class="sbadge ${_badgeCls(r.STATUS_LIQ)}">${r.STATUS_LIQ}</span></td>
        <td class="num">${fmtBRL(r.VALOR_ALVO)}</td><td class="num">${fmtBRL(r.MEDIA_VERBA)}</td>
        <td class="num">${fmtBRL(r.VAR_ABS)}</td><td class="num">${fmtPct(r.VAR_PCT)}</td>
        <td class="num">${fmtBRL(r.LIQUIDO_ALVO)}</td><td>${String(r.AUDITORIA).substring(0,100)}</td>
        ${_confCellFunc(r)}`;
      tr.style.cursor = "pointer";
      sum.view.bind(tr, r);
      tbody.appendChild(tr);
    });
    sum.view.finish(document.getElementById("tabelaColab"));
    _bindConfDelegation("colabBody");
    _bindSort("#tabelaColab", S.colabSort, _renderizarTabelaColab);
    document.getElementById("countColab").textContent = `${sum.page.length} de ${df.length} linhas`;
  }

  // ---- Tabela principal ----
  function _popularCombosTabela() {
    const statuses = [...new Set(S.liquidoAnalise.map(r => r.STATUS))].sort();
    const sel = document.getElementById("comboStatusTab"); sel.innerHTML = "<option>(todos)</option>";
    statuses.forEach(s => { const o = document.createElement("option"); o.value = s; o.textContent = s; sel.appendChild(o); });
  }

  // ---- Bind sort por cabeçalho ----
  let _sortBound = false;
  function _bindSortHeaders() {
    if (_sortBound) return; _sortBound = true;
    // Clique nos th
    document.querySelectorAll("#tabelaPrincipal th[data-col]").forEach(th => {
      th.style.cursor = "pointer";
      th.title = "Clique para ordenar";
      th.addEventListener("click", () => {
        const col = th.dataset.col;
        if (S.sortCol === col) S.sortAsc = !S.sortAsc;
        else { S.sortCol = col; S.sortAsc = false; }
        _renderizarTabela();
      });
    });
    // Combo "Ordenar por" também ordena ao mudar
    document.getElementById("comboOrdemTab").addEventListener("change", () => {
      S.sortCol = document.getElementById("comboOrdemTab").value;
      S.sortAsc = false;
      _renderizarTabela();
    });
  }

  function _limparFiltrosTabela() {
    document.getElementById("comboStatusTab").value = "(todos)";
    document.getElementById("comboOrdemTab").value = "VAR_ABS";
    document.getElementById("entryBusca").value = "";
    const cc = document.getElementById("comboConfTab"); if (cc) cc.value = "todos";
    S.sortCol = "VAR_ABS"; S.sortAsc = false;
    _renderizarTabela();
  }

  function _renderizarTabela() {
    if (!S.liquidoAnalise) return;
    let tab = [...S.liquidoAnalise];
    const stFilt = document.getElementById("comboStatusTab").value;
    if (stFilt && stFilt !== "(todos)") tab = tab.filter(r => r.STATUS === stFilt);
    const busca = (document.getElementById("entryBusca").value || "").toUpperCase().trim();
    if (busca) tab = tab.filter(r => r.Nome.toUpperCase().includes(busca) || String(r.Matrícula).includes(busca));

    // Filtro de conferência
    const fConf = document.getElementById("comboConfTab");
    S.filtroConf = fConf ? fConf.value : "todos";
    if (S.filtroConf === "pendentes") tab = tab.filter(r => !(S.conf[_confKey(r)] && S.conf[_confKey(r)].ok));
    else if (S.filtroConf === "conferidos") tab = tab.filter(r => S.conf[_confKey(r)] && S.conf[_confKey(r)].ok);
    else if (S.filtroConf === "comobs") tab = tab.filter(r => { const c = S.conf[_confKey(r)]; return c && c.obs && c.obs.trim(); });

    // Sincroniza combo com sort atual
    const col = S.sortCol || "VAR_ABS";
    const combo = document.getElementById("comboOrdemTab");
    if ([...combo.options].some(o => o.value === col)) combo.value = col;

    // Ordena — numérico por valor absoluto para impacto, aritmético para outros
    const numCols = new Set(["MEDIA_BASELINE","LIQUIDO_ALVO","VAR_ABS","VAR_PCT","Z_SCORE"]);
    const matCols = new Set(["Matrícula"]);
    tab.sort((a, b) => {
      const va = a[col], vb = b[col];
      if (matCols.has(col)) {
        // Matrícula: tenta numérico, senão string
        const na = parseFloat(va), nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) return S.sortAsc ? na - nb : nb - na;
        return S.sortAsc ? String(va||"").localeCompare(String(vb||""),"pt-BR") : String(vb||"").localeCompare(String(va||""),"pt-BR");
      }
      if (numCols.has(col)) {
        if (col === "VAR_ABS" || col === "VAR_PCT") {
          // Desc padrão = maior impacto absoluto primeiro
          if (!S.sortAsc) return Math.abs(vb || 0) - Math.abs(va || 0);
          // Asc = menor valor primeiro
          const na2 = va ?? Infinity, nb2 = vb ?? Infinity;
          return na2 - nb2;
        }
        const na = va ?? (S.sortAsc ? Infinity : -Infinity);
        const nb = vb ?? (S.sortAsc ? Infinity : -Infinity);
        return S.sortAsc ? na - nb : nb - na;
      }
      return S.sortAsc
        ? String(va||"").localeCompare(String(vb||""),"pt-BR")
        : String(vb||"").localeCompare(String(va||""),"pt-BR");
    });

    // Atualiza indicadores visuais nos cabeçalhos
    document.querySelectorAll("#tabelaPrincipal th").forEach(th => {
      th.classList.remove("sort-asc", "sort-desc");
      if (th.dataset.col === col) th.classList.add(S.sortAsc ? "sort-asc" : "sort-desc");
    });

    const tbody = document.getElementById("tabelaBody"); tbody.innerHTML = "";
    const sum = _sumView('principal', tab, {render: _renderizarTabela, investigate: (r, tr) => _selecionarFunc(r, tr)});
    sum.page.forEach(r => {
      const tr = document.createElement("tr");
      if (STATUS_CRITICOS.has(r.STATUS)) tr.className = "critico";
      else if (r.STATUS === "DESLIGADO") tr.className = "desligado";
      const dNome = maskNome(r.Nome, r.Matrícula);
      const dMat  = maskMat(r.Matrícula);
      tr.innerHTML = `<td class="center">${dMat}</td><td>${String(dNome).substring(0,55)}</td>
        <td class="num">${fmtBRL(r.MEDIA_BASELINE)}</td><td class="num">${fmtBRL(r.LIQUIDO_ALVO)}</td>
        <td class="num">${fmtBRL(r.VAR_ABS)}</td><td class="num">${fmtPct(r.VAR_PCT)}</td>
        <td class="num">${r.Z_SCORE != null ? r.Z_SCORE.toFixed(2) : "-"}</td>
        <td class="status"><span class="sbadge ${_badgeCls(r.STATUS)}">${r.STATUS}</span></td>
        <td>${String(r.AUDITORIA).substring(0,115)}</td>
        ${_confCellFunc(r)}`;
      sum.view.bind(tr, r);
      tbody.appendChild(tr);
    });
    sum.view.finish(document.getElementById("tabelaPrincipal"));
    _bindConfDelegation("tabelaBody");
    _atualizarConfContador();
    document.getElementById("countTabela").textContent =
      `${sum.page.length} de ${tab.length} linhas exibidas (total na análise: ${S.liquido.length})`;
  }

  function _selecionarFunc(r, tr) {
    S.funcSelecionado = r;

    _renderDrill(r);
  }

  function _renderDrill(r) {
    const dNome = maskNome(r.Nome, r.Matrícula);
    const dMat  = maskMat(r.Matrícula);
    document.getElementById("drillInfo").textContent =
      `Mat. ${dMat} — ${dNome}  |  Status: ${r.STATUS}  |  ` +
      `Baseline: ${fmtBRL(r.MEDIA_BASELINE)}  |  Líquido ${S.mesAlvo}: ${fmtBRL(r.LIQUIDO_ALVO)} (${fmtPct(r.VAR_PCT)})  |  ${r.AUDITORIA}`;
    renderDrillFuncionario("chartDrill", r, S.mesesDetectados, S.mesAlvo, r.MEDIA_BASELINE);
    const funcRows = S.dfFiltrado.filter(row => funcKey(row) === funcKey(r));
    const crs   = [...new Set(funcRows.map(row => row["C.R."] || "").filter(Boolean))].sort();
    const clss  = [...new Set(funcRows.map(row => row["Clas."] || "").filter(Boolean))].sort();
    const procs = [...new Set(funcRows.map(row => row["Processo"] || "").filter(Boolean))].sort();
    _setOptions("comboDrillCR", ["(todos)", ...crs]);
    _setOptions("comboDrillClasse", ["(todos)", ...clss]);
    _setOptions("comboDrillProc", ["(todos)", ...procs]);
    document.getElementById("entryDrillVerba").value = "";
    _renderConfBox(r);
    const vh = document.getElementById("btnVerHolerite");
    if (vh) { vh.style.display = "inline-block"; vh.onclick = () => _verHolerite(r.Matrícula, r.Nome, r.Empresa); }
    _renderizarDrillVerbas();
  }

  function _limparFiltrosDrill() {
    ["comboDrillCR","comboDrillClasse","comboDrillProc"].forEach(id => document.getElementById(id).value = "(todos)");
    document.getElementById("entryDrillVerba").value = "";
    _renderizarDrillVerbas();
  }

  function _renderizarDrillVerbas() {
    const tbody = document.getElementById("drillBody"); tbody.innerHTML = "";
    if (!S.funcSelecionado) return;
    const r = S.funcSelecionado;
    let rows = S.dfFiltrado.filter(row => funcKey(row) === funcKey(r));
    const colAlvo = `${S.mesAlvo} - Valor`;
    // Zero é um lançamento selecionável; não elimina contribuição nula.
    // Meses anteriores (3)
    const idxAlvo = S.mesesDetectados.indexOf(S.mesAlvo);
    const meses3 = S.mesesDetectados.slice(Math.max(0, idxAlvo - 3), idxAlvo);
    const nomesM = Array.from({length:3}, (_, i) => meses3[i - (3 - meses3.length)] || "-");
    ["thM3","thM2","thM1"].forEach((id, i) => document.getElementById(id).textContent = nomesM[i]);
    document.getElementById("thAlvo").textContent = S.mesAlvo;
    // Filtros
    const fCR    = document.getElementById("comboDrillCR").value;
    const fClasse = document.getElementById("comboDrillClasse").value;
    const fProc  = document.getElementById("comboDrillProc").value;
    const busca  = (document.getElementById("entryDrillVerba").value || "").toUpperCase();
    if (fCR    && fCR    !== "(todos)") rows = rows.filter(row => (row["C.R."] || "") === fCR);
    if (fClasse && fClasse !== "(todos)") rows = rows.filter(row => row["Clas."] === fClasse);
    if (fProc  && fProc  !== "(todos)") rows = rows.filter(row => row["Processo"] === fProc);
    if (busca) rows = rows.filter(row => String(row["Código"]).toUpperCase().includes(busca) || String(row["Descrição"]).toUpperCase().includes(busca));
    const ordemClas = { PGTO: 0, DESC: 1, OUTRO: 2 };
    if (S.drillSort.key) {
      const mAt = colAlvo;
      const dVal = (row, key) => {
        switch (key) {
          case "cod":  return String(row["Código"]);
          case "desc": return String(row["Descrição"]);
          case "clas": return String(row["Clas."] || "");
          case "cr":   return String(row["C.R."] || "");
          case "proc": return String(row["Processo"] || "");
          case "alvo": return row[mAt] || 0;
          case "m3": case "m2": case "m1": {
            const m = meses3[({ m3: 0, m2: 1, m1: 2 })[key] - (3 - meses3.length)];
            return m ? (row[`${m} - Valor`] || 0) : 0;
          }
          case "conf": { const c = _confGet(_confKey(r)); return (c.linhas && c.linhas[_lineKey(row)]) ? 1 : 0; }
          default: return "";
        }
      };
      rows.sort(_genCmp(dVal, S.drillSort.key, S.drillSort.dir));
    } else {
      rows.sort((a, b) => {
        const oa = ordemClas[a["Clas."]] ?? 99, ob = ordemClas[b["Clas."]] ?? 99;
        return oa !== ob ? oa - ob : String(a["Código"]).localeCompare(String(b["Código"]));
      });
    }
    const sum = _sumView('verbas-detalhe', rows, {
      key: row => JSON.stringify([RowSelection.sourceId(row), S.mesAlvo]), values: _sumPayroll, emptyGroup: 'Folha (proventos − descontos)',
      render: _renderizarDrillVerbas, investigate: row => RowSelection.openDetail('Lançamento de folha · ' + S.mesAlvo, row)
    });
    sum.page.forEach(row => {
      const tr = document.createElement("tr");
      const cl = String(row["Clas."] || "").trim().toUpperCase();
      tr.className = cl === "PGTO" ? "row-pgto" : cl === "DESC" ? "row-desc" : "row-outro";
      const valsM = Array.from({length:3}, (_, i) => {
        const m = meses3[i - (3 - meses3.length)];
        return m ? fmtBRL(row[`${m} - Valor`] || 0) : "-";
      });
      tr.innerHTML = `<td class="center">${row["Código"]}</td><td>${String(row["Descrição"]).substring(0,48)}</td>
        <td class="center">${cl}</td><td class="center">${row["C.R."] || "-"}</td>
        <td>${String(row["Processo"] || "-").substring(0,22)}</td>
        ${valsM.map(v => `<td class="num">${v}</td>`).join("")}
        <td class="num drillAlvo">${fmtBRL(row[colAlvo])}</td>
        ${_confCellLinha(r, row)}`;
      sum.view.bind(tr, row); tbody.appendChild(tr);
    });
    sum.view.finish(document.getElementById("tabelaDrill"));
    _bindConfLinhaDelegation();
    _bindSort("#tabelaDrill", S.drillSort, _renderizarDrillVerbas);
  }

  // ---- Modal Filtrar Verbas ----
  function _abrirFiltroVerbas() {
    document.getElementById("modalFiltroVerbas")?.remove();
    const dfF = S.df.filter(r => S.empresasSel.includes((r["Empresa"] || "").trim()));
    const valCols = S.mesesDetectados.map(m => `${m} - Valor`);
    const agg = {};
    for (const r of dfF) {
      const cod = String(r["Código"]).trim();
      if (!agg[cod]) agg[cod] = { cod, desc: r["Descrição"], clas: r["Clas."], total: 0 };
      for (const c of valCols) agg[cod].total += Math.abs(r[c] || 0);
    }
    const lista = Object.values(agg).sort((a, b) => b.total - a.total);

    const modal = document.createElement("div");
    modal.id = "modalFiltroVerbas";
    modal.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;z-index:1000";
    modal.innerHTML = `
      <div style="background:white;border-radius:6px;width:880px;max-width:95vw;height:78vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,.3)">
        <div style="background:${IG_AZUL};color:white;padding:13px 18px;font-weight:700;font-size:14px;display:flex;justify-content:space-between;align-items:center">
          Gerenciar verbas incluídas no cálculo de outliers
          <button onclick="document.getElementById('modalFiltroVerbas').remove()" style="background:none;border:none;color:white;font-size:22px;cursor:pointer;line-height:1">×</button>
        </div>
        <div style="padding:9px 14px;border-bottom:1px solid #ccc;display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <input id="mvBusca" type="text" placeholder="buscar código ou descrição" class="input-ctrl-sm" style="width:240px">
          <button class="btn-sm btn-primary" onclick="APP._mvFiltrar()">Buscar</button>
          <span style="color:#ccc">|</span>
          <button class="btn-sm btn-secondary-sm" onclick="APP._mvIncluirTodos()">Incluir todos</button>
          <button class="btn-sm btn-secondary-sm" onclick="APP._mvExcluirTodos()">Excluir todos</button>
          <span style="font-size:10px;color:#999">Verba 0020 sempre protegida</span>
        </div>
        <div style="flex:1;overflow:auto">
          <table class="data-table" id="mvTable">
            <thead><tr><th>Inc?</th><th>Código</th><th>Descrição</th><th>Clas.</th><th>Total período</th></tr></thead>
            <tbody id="mvBody"></tbody>
          </table>
        </div>
        <div style="padding:11px 14px;border-top:1px solid #ccc;display:flex;justify-content:space-between;align-items:center">
          <div id="mvStatus" style="font-size:12px;color:#646464"></div>
          <div style="display:flex;gap:8px">
            <button class="btn-sm btn-secondary-sm" onclick="document.getElementById('modalFiltroVerbas').remove()">Cancelar</button>
            <button class="btn-sm btn-primary" onclick="APP._mvAplicar()">Aplicar e fechar</button>
          </div>
        </div>
      </div>`;
    modal._excluidos = new Set(S.verbasExcluidas);
    modal._lista = lista;
    document.body.appendChild(modal);
    setTimeout(() => _mvPopular(), 0);
  }

  function _mvPopular() {
    const modal = document.getElementById("modalFiltroVerbas"); if (!modal) return;
    const busca = (document.getElementById("mvBusca")?.value || "").toUpperCase().trim();
    const tbody = document.getElementById("mvBody"); if (!tbody) return;
    tbody.innerHTML = "";
    let df = modal._lista;
    if (busca) df = df.filter(r => String(r.cod).toUpperCase().includes(busca) || String(r.desc).toUpperCase().includes(busca));
    df.forEach(r => {
      const protegida = normCod(r.cod) === "20";
      const excluida  = !protegida && modal._excluidos.has(r.cod);
      const tr = document.createElement("tr");
      tr.style.cssText = `background:${protegida ? "#E8F0F7" : excluida ? "#FBE8F1" : "white"};${excluida ? `color:${IG_VINHO}` : ""}`;
      tr.innerHTML = `<td class="center" style="font-weight:700">${protegida ? "—" : excluida ? "Não" : "Sim"}</td>
        <td class="center">${r.cod}</td><td>${String(r.desc).substring(0,68)}</td>
        <td class="center">${r.clas}</td><td class="num">${fmtBRL(r.total)}</td>`;
      if (!protegida) {
        tr.style.cursor = "pointer";
        tr.addEventListener("click", () => {
          if (modal._excluidos.has(r.cod)) modal._excluidos.delete(r.cod); else modal._excluidos.add(r.cod);
          _mvPopular();
        });
      }
      tbody.appendChild(tr);
    });
    const statusEl = document.getElementById("mvStatus"); if (!statusEl) return;
    const n = modal._excluidos.size;
    statusEl.textContent = n ? `${n} de ${modal._lista.length} verbas excluídas` : `Todas as ${modal._lista.length} verbas incluídas`;
  }

  function _mvFiltrar() { _mvPopular(); }
  function _mvIncluirTodos() {
    const m = document.getElementById("modalFiltroVerbas"); if (!m) return;
    m._excluidos = new Set(); _mvPopular();
  }
  function _mvExcluirTodos() {
    const m = document.getElementById("modalFiltroVerbas"); if (!m) return;
    m._lista.forEach(r => { if (normCod(r.cod) !== "20") m._excluidos.add(r.cod); }); _mvPopular();
  }
  function _mvAplicar() {
    const m = document.getElementById("modalFiltroVerbas"); if (!m) return;
    S.verbasExcluidas = new Set(m._excluidos); m.remove(); acaoReprocessar();
  }

  // ---- Download modelo CSV ----
  function downloadModeloCSV() {
    const months = ["SET/25","OUT/25","NOV/25","DEZ/25","JAN/26","FEV/26","MAR/26","ABR/26","MAI/26"];
    const mCols  = months.map(m => `${m} - Valor`);
    const sep    = ";";
    const hdrs   = ["Empresa","Processo","Matr\xedcula","Nome","Data Rescis\xe3o",
                    "Clas.","C\xf3digo","Descri\xe7\xe3o","C.R.",...mCols];

    const br = v => {
      if (!v) return "0,00";
      const neg = v < 0; v = Math.abs(v);
      const [ip, dp] = v.toFixed(2).split(".");
      let r = ""; [...ip].reverse().forEach((c,i) => { if (i && i%3===0) r="."+r; r=c+r; });
      return (neg ? "-" : "") + r + "," + dp;
    };
    const row = (emp,proc,mat,nome,resc,clas,cod,desc,cr,...vals) =>
      [emp,proc,mat,nome,resc,clas,cod,desc,cr,...vals.map(br)].join(sep);

    const lines = [
      hdrs.join(sep),
      // Func 001 — normal
      row("EMPRESA ALPHA S.A.","MENSAL","00001","JOAO DA SILVA SANTOS","","PGTO","0020","SALARIO BASE","CC-ADM",3500,3500,3500,3500,3500,3500,3500,3500,3500),
      row("EMPRESA ALPHA S.A.","MENSAL","00001","JOAO DA SILVA SANTOS","","PGTO","0040","HORA EXTRA 50%","CC-ADM",0,245,0,180,0,0,0,0,0),
      row("EMPRESA ALPHA S.A.","MENSAL","00001","JOAO DA SILVA SANTOS","","PGTO","9950","LIQUIDO MENSAL","CC-ADM",2855,3087,2855,3012,2855,2855,2855,2855,2855),
      row("EMPRESA ALPHA S.A.","MENSAL","00001","JOAO DA SILVA SANTOS","","DESC","3010","INSS","CC-ADM",385,385,385,385,385,385,385,385,385),
      row("EMPRESA ALPHA S.A.","MENSAL","00001","JOAO DA SILVA SANTOS","","DESC","3020","IRRF","CC-ADM",210,252,210,244,210,210,210,210,210),
      row("EMPRESA ALPHA S.A.","MENSAL","00001","JOAO DA SILVA SANTOS","","DESC","3050","PLANO SAUDE","CC-ADM",200,200,200,200,200,200,200,200,200),
      row("EMPRESA ALPHA S.A.","MENSAL","00001","JOAO DA SILVA SANTOS","","DESC","3060","VT","CC-ADM",150,150,150,150,150,150,150,150,150),
      // Func 002 — salário alto
      row("EMPRESA ALPHA S.A.","MENSAL","00002","MARIA COSTA OLIVEIRA","","PGTO","0020","SALARIO BASE","CC-ADM",8000,8000,8000,8000,8000,8000,8000,8000,8000),
      row("EMPRESA ALPHA S.A.","MENSAL","00002","MARIA COSTA OLIVEIRA","","PGTO","9950","LIQUIDO MENSAL","CC-ADM",5520,5520,5520,5520,5520,5520,5520,5520,5520),
      row("EMPRESA ALPHA S.A.","MENSAL","00002","MARIA COSTA OLIVEIRA","","DESC","3010","INSS","CC-ADM",880,880,880,880,880,880,880,880,880),
      row("EMPRESA ALPHA S.A.","MENSAL","00002","MARIA COSTA OLIVEIRA","","DESC","3020","IRRF","CC-ADM",960,960,960,960,960,960,960,960,960),
      row("EMPRESA ALPHA S.A.","MENSAL","00002","MARIA COSTA OLIVEIRA","","DESC","3050","PLANO SAUDE","CC-ADM",640,640,640,640,640,640,640,640,640),
      // Func 003 — OUTLIER em MAI/26
      row("EMPRESA ALPHA S.A.","MENSAL","00003","CARLOS PEREIRA MENDES","","PGTO","0020","SALARIO BASE","CC-VND",4200,4200,4200,4200,4200,4200,4200,4200,4200),
      row("EMPRESA ALPHA S.A.","MENSAL","00003","CARLOS PEREIRA MENDES","","PGTO","0040","HORA EXTRA 50%","CC-VND",0,0,0,0,0,0,0,0,3800),
      row("EMPRESA ALPHA S.A.","MENSAL","00003","CARLOS PEREIRA MENDES","","PGTO","9950","LIQUIDO MENSAL","CC-VND",3020,3020,3020,3020,3020,3020,3020,3020,6480),
      row("EMPRESA ALPHA S.A.","MENSAL","00003","CARLOS PEREIRA MENDES","","DESC","3010","INSS","CC-VND",462,462,462,462,462,462,462,462,462),
      row("EMPRESA ALPHA S.A.","MENSAL","00003","CARLOS PEREIRA MENDES","","DESC","3020","IRRF","CC-VND",618,618,618,618,618,618,618,618,1120),
      row("EMPRESA ALPHA S.A.","MENSAL","00003","CARLOS PEREIRA MENDES","","DESC","3050","PLANO SAUDE","CC-VND",100,100,100,100,100,100,100,100,100),
      // Func 004 — DESLIGADO 15/03/2026
      row("EMPRESA BETA LTDA","MENSAL","00004","ANA PAULA RODRIGUES","15/03/2026","PGTO","0020","SALARIO BASE","CC-FIN",5000,5000,5000,5000,5000,5000,2500,0,0),
      row("EMPRESA BETA LTDA","MENSAL","00004","ANA PAULA RODRIGUES","15/03/2026","PGTO","9950","LIQUIDO MENSAL","CC-FIN",3480,3480,3480,3480,3480,3480,1740,0,0),
      row("EMPRESA BETA LTDA","MENSAL","00004","ANA PAULA RODRIGUES","15/03/2026","DESC","3010","INSS","CC-FIN",550,550,550,550,550,550,275,0,0),
      row("EMPRESA BETA LTDA","MENSAL","00004","ANA PAULA RODRIGUES","15/03/2026","DESC","3020","IRRF","CC-FIN",810,810,810,810,810,810,405,0,0),
      row("EMPRESA BETA LTDA","MENSAL","00004","ANA PAULA RODRIGUES","15/03/2026","DESC","3060","VT","CC-FIN",160,160,160,160,160,160,80,0,0),
      // Func 005 — NOVO (sem baseline)
      row("EMPRESA BETA LTDA","MENSAL","00005","LUCAS FERREIRA LIMA","","PGTO","0020","SALARIO BASE","CC-FIN",0,0,0,0,0,0,0,0,3200),
      row("EMPRESA BETA LTDA","MENSAL","00005","LUCAS FERREIRA LIMA","","PGTO","9950","LIQUIDO MENSAL","CC-FIN",0,0,0,0,0,0,0,0,2260),
      row("EMPRESA BETA LTDA","MENSAL","00005","LUCAS FERREIRA LIMA","","DESC","3010","INSS","CC-FIN",0,0,0,0,0,0,0,0,352),
      row("EMPRESA BETA LTDA","MENSAL","00005","LUCAS FERREIRA LIMA","","DESC","3060","VT","CC-FIN",0,0,0,0,0,0,0,0,150),
      // Func 006 — FERIAS
      row("EMPRESA BETA LTDA","FERIAS","00006","SANDRA MOURA COSTA","","PGTO","0020","SALARIO BASE","CC-RH",0,6000,0,0,0,0,0,6000,0),
      row("EMPRESA BETA LTDA","FERIAS","00006","SANDRA MOURA COSTA","","PGTO","0260","FERIAS","CC-RH",0,6000,0,0,0,0,0,6000,0),
      row("EMPRESA BETA LTDA","FERIAS","00006","SANDRA MOURA COSTA","","PGTO","0270","1/3 FERIAS","CC-RH",0,2000,0,0,0,0,0,2000,0),
      row("EMPRESA BETA LTDA","FERIAS","00006","SANDRA MOURA COSTA","","DESC","3010","INSS","CC-RH",0,660,0,0,0,0,0,660,0),
      row("EMPRESA BETA LTDA","FERIAS","00006","SANDRA MOURA COSTA","","DESC","3020","IRRF","CC-RH",0,480,0,0,0,0,0,480,0),
    ];

    // Codifica como ISO-8859-1 (os headers têm acentos via escape \xNN)
    const content = lines.join("\n");
    const bytes = new Uint8Array(content.length);
    for (let i = 0; i < content.length; i++) bytes[i] = content.charCodeAt(i) & 0xFF;
    const blob = new Blob([bytes], { type: "text/csv;charset=iso-8859-1" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = "modelo_folha_auditoria.csv";
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(a.href);
  }
  function acaoExportarExcel() {
    if (!S.liquido) return;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(S.resumo), "Resumo_Mensal");
    const ordemStatus = ["AUSENTE","NEGATIVO","ZERO_SUSPEITO","EXTREMA","EXTREMA_IQR","QUEDA_LIQUIDO","Z_2SIGMA","ALTA","ALTA_IQR","MATERIAL","NOVO_FUNC","SEM_HISTORICO","OK","SEM_DADOS","DESLIGADO"];
    const liqExp = S.liquido.map(r => {
      const c = S.conf[_confKeyMN(r.Matrícula, r.Nome)] || {};
      return {
      Empresa: r.Empresa, Matrícula: r.Matrícula, Nome: r.Nome,
      MEDIA_BASELINE: r.MEDIA_BASELINE, LIQUIDO_ALVO: r.LIQUIDO_ALVO,
      VAR_ABS: r.VAR_ABS, VAR_PCT: r.VAR_PCT, Z_SCORE: r.Z_SCORE,
      V9950_ALVO: r.V9950_ALVO, DIFF_9950: r.DIFF_9950, FLAG_9950: r.FLAG_9950 ? "SIM" : "",
      STATUS: r.STATUS, AUDITORIA: r.AUDITORIA,
      CONFERIDO: c.ok ? "SIM" : "",
      VERBAS_CONFERIDAS: c.linhas ? Object.keys(c.linhas).length : 0,
      OBSERVACAO: c.obs || "",
      ...Object.fromEntries(S.mesesDetectados.map(m => [m, r[m]]))
      };
    }).sort((a, b) => {
      const ia = ordemStatus.indexOf(a.STATUS), ib = ordemStatus.indexOf(b.STATUS);
      return ia !== ib ? ia - ib : Math.abs(b.VAR_ABS || 0) - Math.abs(a.VAR_ABS || 0);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(liqExp), "Liquido_Funcionario");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(S.zeradosPGTO), "Verbas_PGTO_Zeradas");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(S.zeradosDESC), "Verbas_DESC_Zeradas");
    // Aba de conferência (revisão humana)
    const confExp = [];
    (S.liquidoAnalise || S.liquido).forEach(r => {
      const c = S.conf[_confKeyMN(r.Matrícula, r.Nome)];
      if (!c) return;
      const nLin = c.linhas ? Object.keys(c.linhas).length : 0;
      if (!c.ok && !(c.obs && c.obs.trim()) && !nLin) return;
      confExp.push({
        Matrícula: r.Matrícula, Nome: r.Nome, STATUS: r.STATUS,
        CONFERIDO: c.ok ? "SIM" : "",
        VERBAS_CONFERIDAS: nLin,
        OBSERVACAO: c.obs || "",
        ATUALIZADO_EM: c.ts ? new Date(c.ts).toLocaleString("pt-BR") : ""
      });
    });
    if (confExp.length)
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(confExp), "Conferencia");
    // Aba de descontos para revisar
    const descExp = (S.descontosRevisar || []).map(e => ({
      Empresa: e.Empresa, Matrícula: e.Matrícula, Nome: e.Nome,
      Código: e.Código, Descrição: e.Descrição, Valor: e.VALOR, Referencia: e.REFERENCIA,
      Tipo: e.TIPO, Prioridade: e.PRIORIDADE,
      Comprometimento_pct: e.COMPROMETIMENTO != null ? +e.COMPROMETIMENTO.toFixed(2) : "",
      Z_verba: e.Z_VERBA != null ? +e.Z_VERBA.toFixed(2) : "",
      Motivo: e.MOTIVO
    }));
    if (descExp.length)
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(descExp), "Descontos_Revisar");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
      { chave: "metodo", valor: S.metadata.metodo },
      { chave: "config", valor: S.configNome },
      { chave: "mes_alvo", valor: S.mesAlvo },
      { chave: "meses_baseline", valor: S.mesesBaseline.join(", ") },
      { chave: "folha_total_alvo", valor: S.metadata.folhaTotalAlvo },
      { chave: "materialidade", valor: S.metadata.materialidade },
      { chave: "Q1_var_abs", valor: S.metadata.q1 },
      { chave: "Q3_var_abs", valor: S.metadata.q3 },
      { chave: "IQR_var_abs", valor: S.metadata.iqr },
      { chave: "gerado_em", valor: new Date().toLocaleString("pt-BR") }
    ]), "Parametros");
    XLSX.writeFile(wb, `auditoria_folha_${S.mesAlvo.replace("/","-")}.xlsx`);
    const el = document.getElementById("exportStatus");
    el.textContent = `Relatório gerado: ${liqExp.length} funcionários, ${S.mesesDetectados.length} meses.`;
    el.style.color = IG_VERDE_ESC;
  }

  // ---- Helpers ----
  function _card(id, titulo, valor, delta, cls) {
    const el = document.getElementById(id); if (!el) return;
    el.className = "metric-card" + (cls === "alerta" ? " alerta" : cls === "atencao" ? " atencao" : "");
    el.innerHTML = `<div class="card-titulo">${titulo}</div>
      <div class="card-valor">${valor}</div>
      <div class="card-delta ${(cls==="alerta"||cls==="atencao") ? "negativo" : ""}">${delta}</div>`;
  }
  function _badgeCls(s) {
    if (STATUS_CRITICOS.has(s)) return "sbadge-critico";
    if (s === "DESLIGADO") return "sbadge-desligado";
    if (s === "NOVO_FUNC") return "sbadge-novo";
    return "sbadge-ok";
  }
  function _setOptions(id, opts) {
    const sel = document.getElementById(id); sel.innerHTML = "";
    opts.forEach(v => { const o = document.createElement("option"); o.value = v; o.textContent = v; sel.appendChild(o); });
  }

  // ---- Public API ----
  return { init, acaoReprocessar, acaoExportarExcel, downloadModeloCSV, acaoResetDados,
    selecionarClasseVerbas, limparSelecaoVerbas,
    _mvFiltrar, _mvIncluirTodos, _mvExcluirTodos, _mvAplicar,
    _selTodosEmpresa, _selTodosProcesso, _navigateToFuncDrill,
    confExport, confImport, confReset, _verHolerite };
})();

document.addEventListener("DOMContentLoaded", APP.init);
