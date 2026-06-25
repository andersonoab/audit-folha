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
    // análise de verbas
    verbasAvCompleta: null, dfAvBase: null,
    codigosColabAtual: [], colabDfAtual: null,
    topVerbasDF: null,
    // tabela
    funcSelecionado: null,
    sortCol: "VAR_ABS", sortAsc: false,
    tabsDirty: {},
    // privacidade
    anonimizar: false,
    mapaAnonimo: {},
    csvBruto: null, csvNome: null,
    // conferência (revisão humana)
    conf: {}, filtroConf: "todos"
  };

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
  const CONF_KEY = "af_conf_v1";
  function _confKey(r)   { return String(r.Matrícula) + "|||" + String(r.Nome); }
  function _confKeyMN(m, n) { return String(m) + "|||" + String(n); }
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
    // Pendentes primeiro; dentro de cada grupo, maior impacto
    list.sort((a, b) => {
      const ca = (S.conf[_confKey(a)] && S.conf[_confKey(a)].ok) ? 1 : 0;
      const cb = (S.conf[_confKey(b)] && S.conf[_confKey(b)].ok) ? 1 : 0;
      if (ca !== cb) return ca - cb;
      return Math.abs(b.VAR_ABS || 0) - Math.abs(a.VAR_ABS || 0);
    });
    list.slice(0, 2000).forEach(r => {
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
        <td class="center"><button class="btn-xs conf-open" data-mat="${escAttr(r.Matrícula)}" data-nome="${escAttr(r.Nome)}">Abrir</button></td>`;
      tbody.appendChild(tr);
    });
    _bindConfTab();
    document.getElementById("countConf").textContent = `${Math.min(2000, list.length)} de ${list.length} colaboradores`;
    _atualizarConfContador();
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
      const tr = chk.closest("tr"); if (tr) tr.className = chk.checked ? "conf-row-ok" : "";
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
      _navigateToFuncDrill(btn.dataset.mat, btn.dataset.nome);
    });
  }
  function _confMarcarVisiveis() {
    const cells = document.querySelectorAll("#confBody td.conf-cell[data-confkey]");
    if (!cells.length) return;
    if (!confirm(`Marcar ${cells.length} colaborador(es) visível(eis) como conferido(s)?`)) return;
    cells.forEach(cell => _confSetFunc(cell.dataset.confkey, true));
    _renderConferencia(); _renderizarTabela(); _repintarColabConf();
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
    if (!S.empresasSel.length) { alert("Selecione ao menos uma empresa."); return; }
    if (!S.mesesBaseline.length) { alert("Selecione ao menos um mês de baseline."); return; }

    let dfF = S.df.filter(r => S.empresasSel.includes((r["Empresa"] || "").trim()));
    if (S.processosDisp.length > 1) dfF = dfF.filter(r => S.processosSel.includes((r["Processo"] || "").trim()));
    if (!dfF.length) { alert("Nenhuma linha após os filtros de empresa/processo."); return; }

    if (S.verbasExcluidas.size > 0) {
      const exclNorm = new Set([...S.verbasExcluidas].map(normCod));
      exclNorm.delete("20"); // verba 0020 sempre protegida
      dfF = dfF.filter(r => !exclNorm.has(normCod(r["Código"])));
    }
    S.dfFiltrado = dfF;

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
    const { result: liq, metadata } = liquidoPorFuncionario(dfF, S.mesesDetectados, S.mesAlvo, S.mesesBaseline, desligados, cfg);
    S.liquido    = liq;
    S.metadata   = metadata;
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
    S.funcSelecionado = null;
    document.getElementById("drillInfo").textContent = "(selecione um funcionário na tabela acima)";
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
    const _med = (fn) => lb.length ? mean(lb.map(fn)) : 0;
    const mPGTO = _med(r=>r.PGTO), mDESC = _med(r=>r.DESC), mLiq = _med(r=>r.Liquido);
    const vP = mPGTO ? ((la.PGTO/mPGTO)-1)*100 : 0;
    const vD = mDESC ? ((la.DESC/mDESC)-1)*100 : 0;
    const vL = mLiq  ? ((la.Liquido/mLiq)-1)*100 : 0;
    const diff9950 = la.Liquido - la.Verba9950;
    const pct9950  = la.Verba9950 ? (diff9950/la.Verba9950)*100 : 0;
    _card("cardPGTO", `PGTO ${S.mesAlvo}`, fmtBRL(la.PGTO),   `${fmtPct(vP)} vs baseline`, vP < -10 ? "alerta" : "");
    _card("cardDESC", `DESC ${S.mesAlvo}`, fmtBRL(la.DESC),   `${fmtPct(vD)} vs baseline`, vD < -10 ? "alerta" : "");
    _card("cardLIQ",  `Líquido ${S.mesAlvo}`, fmtBRL(la.Liquido), `${fmtPct(vL)} vs baseline`, vL < -10 ? "alerta" : "");
    _card("card9950", "Verba 9950", fmtBRL(la.Verba9950), `Diff: ${fmtBRL(diff9950)} (${fmtPct(pct9950)})`, Math.abs(pct9950) > 15 ? "alerta" : "");
    const nDesl = S.liquido.filter(r => r.STATUS === "DESLIGADO").length;
    document.getElementById("avisoDesligados").textContent = nDesl
      ? `Desligados detectados: ${nDesl} | ${S.excluirDesligados ? "excluídos da análise de líquido" : "incluídos"}.` : "";
    renderEvolucaoMensal("chartEvolucao", S.resumo, S.mesAlvo, S.mesesBaseline);
    renderConciliacao("chartConciliacao", la, S.mesAlvo);
    renderDistribuicaoStatus("chartStatus", S.liquidoAnalise);
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

  // ---- Navegação: scatter → Tabela & Drill-down ----
  function _navigateToFuncDrill(mat, nome) {
    _switchTab("tabela");
    const r = (S.liquido || []).find(x => x.Matrícula === mat && x.Nome === nome);
    if (!r) return;
    // Limpa filtros para garantir que o funcionário apareça
    document.getElementById("comboStatusTab").value = "(todos)";
    const cc = document.getElementById("comboConfTab"); if (cc) cc.value = "todos";
    document.getElementById("entryBusca").value = String(mat);
    S.sortCol = "VAR_ABS"; S.sortAsc = false;
    _renderizarTabela();
    requestAnimationFrame(() => {
      const rows = document.querySelectorAll("#tabelaBody tr");
      for (const tr of rows) {
        const tdMat = tr.querySelector("td.center");
        if (tdMat && tdMat.textContent.trim() === String(mat)) {
          document.querySelectorAll("#tabelaBody tr").forEach(t => t.classList.remove("selected"));
          tr.classList.add("selected");
          tr.scrollIntoView({ behavior: "smooth", block: "center" });
          _renderDrill(r);
          break;
        }
      }
    });
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
      getMatricula(row) === r.Matrícula && getNome(row) === r.Nome
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
        _navigateToFuncDrill(r.Matrícula, r.Nome);
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
    if (lb.options.length > 0 && lb.selectedIndex < 0) { lb.selectedIndex = 0; _renderDrillVerba(); }
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
    if (!cods.length) { infoEl.textContent = "(nenhuma selecionada)"; _limparCardsAV(); return; }
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
    if (!cods.length) { S.colabDfAtual = null; document.getElementById("colabBody").innerHTML = ""; document.getElementById("countColab").textContent = ""; return; }
    const alvo = new Set(cods.map(normCod));
    const sub = S.dfAvBase.filter(r => alvo.has(normCod(r["Código"])));
    const funcMap = {};
    for (const r of sub) {
      const key = getMatricula(r) + "|||" + getNome(r);
      if (!funcMap[key]) funcMap[key] = { mat: getMatricula(r), nome: getNome(r), vAlvo: 0, totalPorMes: {} };
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
      const liqRow = S.liquido ? S.liquido.find(r => r.Matrícula === d.mat && r.Nome === d.nome) : null;
      colab.push({ Matrícula: d.mat, Nome: d.nome, STATUS_LIQ: liqRow?.STATUS || "-",
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

  function _renderizarTabelaColab() {
    if (!S.colabDfAtual) return;
    let df = [...S.colabDfAtual];
    const stFilt = document.getElementById("comboColabStatus").value;
    if (stFilt && stFilt !== "(todos)") df = df.filter(r => r.STATUS_LIQ === stFilt);
    const busca = (document.getElementById("entryColabBusca").value || "").toUpperCase().trim();
    if (busca) df = df.filter(r => r.Nome.toUpperCase().includes(busca) || String(r.Matrícula).includes(busca));
    df.sort((a, b) => Math.abs(b.VAR_ABS || 0) - Math.abs(a.VAR_ABS || 0));
    const tbody = document.getElementById("colabBody"); tbody.innerHTML = "";
    df.slice(0,2000).forEach(r => {
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
      tr.addEventListener("click", () => {
        document.querySelectorAll("#colabBody tr").forEach(t => t.classList.remove("selected"));
        tr.classList.add("selected");
        _renderColabDrill(r);
        document.getElementById("colabDrillInfo")?.scrollIntoView({ behavior: "smooth", block: "nearest" });
      });
      tbody.appendChild(tr);
    });
    _bindConfDelegation("colabBody");
    document.getElementById("countColab").textContent = `${Math.min(2000,df.length)} de ${df.length} linhas`;
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
    tab.slice(0,1000).forEach(r => {
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
      tr.addEventListener("click", () => _selecionarFunc(r, tr));
      tbody.appendChild(tr);
    });
    _bindConfDelegation("tabelaBody");
    _atualizarConfContador();
    document.getElementById("countTabela").textContent =
      `${Math.min(1000,tab.length)} de ${tab.length} linhas exibidas (total na análise: ${S.liquido.length})`;
  }

  function _selecionarFunc(r, tr) {
    S.funcSelecionado = r;
    document.querySelectorAll("#tabelaBody tr").forEach(t => t.classList.remove("selected"));
    tr.classList.add("selected");
    _renderDrill(r);
  }

  function _renderDrill(r) {
    const dNome = maskNome(r.Nome, r.Matrícula);
    const dMat  = maskMat(r.Matrícula);
    document.getElementById("drillInfo").textContent =
      `Mat. ${dMat} — ${dNome}  |  Status: ${r.STATUS}  |  ` +
      `Baseline: ${fmtBRL(r.MEDIA_BASELINE)}  |  Líquido ${S.mesAlvo}: ${fmtBRL(r.LIQUIDO_ALVO)} (${fmtPct(r.VAR_PCT)})  |  ${r.AUDITORIA}`;
    renderDrillFuncionario("chartDrill", r, S.mesesDetectados, S.mesAlvo, r.MEDIA_BASELINE);
    const funcRows = S.dfFiltrado.filter(row => getMatricula(row) === r.Matrícula && getNome(row) === r.Nome);
    const crs   = [...new Set(funcRows.map(row => row["C.R."] || "").filter(Boolean))].sort();
    const clss  = [...new Set(funcRows.map(row => row["Clas."] || "").filter(Boolean))].sort();
    const procs = [...new Set(funcRows.map(row => row["Processo"] || "").filter(Boolean))].sort();
    _setOptions("comboDrillCR", ["(todos)", ...crs]);
    _setOptions("comboDrillClasse", ["(todos)", ...clss]);
    _setOptions("comboDrillProc", ["(todos)", ...procs]);
    document.getElementById("entryDrillVerba").value = "";
    _renderConfBox(r);
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
    let rows = S.dfFiltrado.filter(row => getMatricula(row) === r.Matrícula && getNome(row) === r.Nome);
    const colAlvo = `${S.mesAlvo} - Valor`;
    rows = rows.filter(row => Math.abs(row[colAlvo] || 0) > 0.001);
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
    rows.sort((a, b) => {
      const oa = ordemClas[a["Clas."]] ?? 99, ob = ordemClas[b["Clas."]] ?? 99;
      return oa !== ob ? oa - ob : String(a["Código"]).localeCompare(String(b["Código"]));
    });
    rows.forEach(row => {
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
      tbody.appendChild(tr);
    });
    _bindConfLinhaDelegation();
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
    const ordemStatus = ["AUSENTE","NEGATIVO","ZERO_SUSPEITO","EXTREMA","EXTREMA_IQR","Z_2SIGMA","ALTA","ALTA_IQR","MATERIAL","NOVO_FUNC","OK","SEM_DADOS","DESLIGADO"];
    const liqExp = S.liquido.map(r => {
      const c = S.conf[_confKeyMN(r.Matrícula, r.Nome)] || {};
      return {
      Matrícula: r.Matrícula, Nome: r.Nome,
      MEDIA_BASELINE: r.MEDIA_BASELINE, LIQUIDO_ALVO: r.LIQUIDO_ALVO,
      VAR_ABS: r.VAR_ABS, VAR_PCT: r.VAR_PCT, Z_SCORE: r.Z_SCORE,
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
    confExport, confImport, confReset };
})();

document.addEventListener("DOMContentLoaded", APP.init);
