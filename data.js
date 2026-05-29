// ============================================================
// data.js — Auditoria de Folha | Igarapé Digital
// Portado de Python/pandas para JS vanilla
// ============================================================

// ---- Paleta Igarapé Digital ----
const IG_AZUL        = "#0083CA";
const IG_AZUL_ESC    = "#003C64";
const IG_VERDE_ESC   = "#005A64";
const IG_AZUL_CLR    = "#6EB4DC";
const IG_VINHO       = "#7D0041";
const IG_LARANJA     = "#8C321E";
const IG_VERMELHO    = "#7D0041";
const TEXTO_PRINCIPAL  = "#333333";
const TEXTO_SECUNDARIO = "#646464";
const GRID_CINZA     = "#CCCCCC";
const COR_DESLIGADO  = "#646464";

const STATUS_CRITICOS = new Set([
  "AUSENTE","NEGATIVO","ZERO_SUSPEITO","EXTREMA","ALTA","Z_2SIGMA",
  "EXTREMA_IQR","ALTA_IQR","MAD_OUTLIER","MATERIAL"
]);

const PARAMETROS_MERCADO = {
  "Shewhart 3-sigma + Materialidade ISA 320 (recomendado)": {
    metodo: "sigma", alta: 2.0, extrema: 3.0, materialidade_pct: 0.01,
    descricao: "Padrão SPC industrial (Western Electric). |z|>=3 ~ 0,3% esperado. ISA 320: 1% da folha total."
  },
  "Tukey IQR (boxplot clássico)": {
    metodo: "iqr", alta: 1.5, extrema: 3.0, materialidade_pct: 0.01,
    descricao: "Tukey 1977. Robusto a distribuições não-normais. 1,5x IQR = outlier moderado, 3x = extremo."
  },
  "MAD modificado (Iglewicz-Hoaglin)": {
    metodo: "mad", alta: 2.5, extrema: 3.5, materialidade_pct: 0.01,
    descricao: "Mediana absoluta de desvio. Mais robusto que sigma. Threshold 3,5 (paper original 1993)."
  },
  "Anderson Legacy (98%/30%)": {
    metodo: "legacy", alta_pct: 30, alta_abs: 200, extrema_pct: 98, extrema_abs: 500, materialidade_pct: 0.0,
    descricao: "Heurística original Igarapé Digital. Mantida para compatibilidade."
  }
};

// ---- Formatadores ----
function fmtBRL(x) {
  if (x == null || isNaN(x)) return "-";
  const abs = Math.abs(x);
  const parts = abs.toFixed(2).split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  const s = parts[0] + "," + parts[1];
  return x < 0 ? `-R$ ${s}` : `R$ ${s}`;
}
function fmtBRLEixo(x) {
  if (Math.abs(x) >= 1_000_000) return `${(x/1_000_000).toFixed(1)}M`;
  if (Math.abs(x) >= 1_000) return `${(x/1_000).toFixed(0)}k`;
  return x.toFixed(0);
}
function fmtPct(x) {
  if (x == null || isNaN(x)) return "-";
  return `${x >= 0 ? "+" : ""}${x.toFixed(1)}%`;
}
function fmtInt(x) {
  if (x == null || isNaN(x)) return "-";
  return Math.round(x).toLocaleString("pt-BR");
}
function corStatus(status) {
  if (STATUS_CRITICOS.has(status)) return IG_VERMELHO;
  const m = { OK: IG_AZUL_CLR, NOVO_FUNC: IG_AZUL_ESC, SEM_DADOS: GRID_CINZA, DESLIGADO: COR_DESLIGADO };
  return m[status] || GRID_CINZA;
}
function normCod(cod) {
  return String(cod).trim().replace(/^0+/, "") || "0";
}

// ---- Estatísticas ----
function mean(arr) {
  if (!arr.length) return 0;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}
function std(arr) {
  if (arr.length < 2) return 0;
  const m = mean(arr);
  return Math.sqrt(arr.reduce((s, v) => s + (v - m) ** 2, 0) / (arr.length - 1));
}
function median(arr) {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 !== 0 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}
function quantile(sortedArr, q) {
  const pos = (sortedArr.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  return sortedArr[base + 1] !== undefined
    ? sortedArr[base] + rest * (sortedArr[base + 1] - sortedArr[base])
    : sortedArr[base];
}
function calcMAD(vals) {
  if (vals.length < 2) return null;
  const med = median(vals);
  return median(vals.map(v => Math.abs(v - med)));
}

// ---- Mapeamento de meses ----
const MAP_MES = {JAN:1,FEV:2,MAR:3,ABR:4,MAI:5,JUN:6,JUL:7,AGO:8,SET:9,OUT:10,NOV:11,DEZ:12};
function ultimoDiaMes(rotulo) {
  try {
    const [mmm, aa] = rotulo.split("/");
    const mes = MAP_MES[mmm.toUpperCase()];
    let ano = parseInt(aa, 10);
    if (ano < 80) ano += 2000; else ano += 1900;
    return new Date(ano, mes, 0); // day 0 = último dia do mês anterior
  } catch { return null; }
}

// ---- Carregamento do CSV ----
function carregarCSV(text) {
  const parsed = Papa.parse(text, { header: true, delimiter: ";", skipEmptyLines: true });
  let df = parsed.data;
  df = df.map(row => { const r = {...row}; delete r["Unnamed: 49"]; return r; });
  df = df.filter(r => (r["Empresa"] || "").trim() !== "Empresa");
  if (!df.length) return { df: [], meses: [] };

  const headers = Object.keys(df[0]);
  const colsValor = headers.filter(h => h.includes(" - Valor"));
  const meses = colsValor.map(c => c.replace(" - Valor", ""));

  df = df.map(row => {
    const r = {...row};
    for (const col of colsValor) {
      let v = String(r[col] || "0").trim().replace(/\./g, "").replace(",", ".");
      r[col] = parseFloat(v) || 0;
    }
    return r;
  });
  return { df, meses };
}

// ---- Detectar desligados ----
function detectarDesligados(df, mesAlvo) {
  const desligados = {};
  const visto = new Set();
  const limite = ultimoDiaMes(mesAlvo);
  for (const row of df) {
    const mat = row["Matrícula"] || row["Matricula"] || "";
    const nome = row["Nome"] || "";
    const key = mat + "|||" + nome;
    if (visto.has(key)) continue;
    visto.add(key);
    const rescStr = (row["Data Rescisão"] || row["Data Rescisao"] || "").trim();
    if (!rescStr) continue;
    const parts = rescStr.split("/");
    if (parts.length !== 3) continue;
    const dataResc = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    if (isNaN(dataResc.getTime())) continue;
    if (!limite || dataResc <= limite) desligados[key] = dataResc;
  }
  return desligados;
}

// Helper: pega campo matrícula tolerando variações de nome de coluna
function getMatricula(row) { return row["Matrícula"] || row["Matricula"] || ""; }
function getNome(row) { return row["Nome"] || ""; }

// ---- Resumo macro ----
function resumoMacro(df, meses) {
  return meses.map(m => {
    const col = `${m} - Valor`;
    const pgtoRows = df.filter(r => r["Clas."] === "PGTO");
    const descRows  = df.filter(r => r["Clas."] === "DESC");
    const v9950     = df.filter(r => String(r["Código"]).trim() === "9950");
    const pgto = pgtoRows.reduce((s, r) => s + (r[col] || 0), 0);
    const desc = descRows.reduce((s, r) => s + (r[col] || 0), 0);
    const verba9950 = v9950.reduce((s, r) => s + (r[col] || 0), 0);
    const matsAtivos = new Set(pgtoRows.filter(r => (r[col] || 0) !== 0).map(getMatricula));
    return { Mes: m, PGTO: pgto, DESC: desc, Liquido: pgto - desc, Verba9950: verba9950, FuncPGTO: matsAtivos.size };
  });
}

// ---- Salário verba 0020 por mês ----
function salarioVerba20PorMes(df, meses, codigoVerba = "0020") {
  const alvo = normCod(codigoVerba);
  const sub = df.filter(r => normCod(r["Código"]) === alvo);
  return meses.map(m => {
    const col = `${m} - Valor`;
    const ativos = sub.filter(r => (r[col] || 0) > 0);
    const mats = new Set(ativos.map(getMatricula));
    const hc = mats.size;
    const total = ativos.reduce((s, r) => s + (r[col] || 0), 0);
    return { Mes: m, HC: hc, Salario_Total: total, Salario_Medio: hc ? total / hc : 0 };
  });
}

// ---- Líquido por funcionário (núcleo da auditoria) ----
function liquidoPorFuncionario(df, meses, mesAlvo, mesesBaseline, desligados, configParam) {
  configParam = configParam || PARAMETROS_MERCADO["Shewhart 3-sigma + Materialidade ISA 320 (recomendado)"];
  desligados = desligados || {};

  // Índice por mat+nome para performance
  const idxPGTO = {}, idxDESC = {};
  for (const r of df) {
    const key = getMatricula(r) + "|||" + getNome(r);
    const clas = r["Clas."] || "";
    if (clas === "PGTO") { if (!idxPGTO[key]) idxPGTO[key] = []; idxPGTO[key].push(r); }
    if (clas === "DESC") { if (!idxDESC[key]) idxDESC[key] = []; idxDESC[key].push(r); }
  }

  const funcSet = new Map();
  for (const r of df) {
    const key = getMatricula(r) + "|||" + getNome(r);
    if (!funcSet.has(key)) funcSet.set(key, { mat: getMatricula(r), nome: getNome(r) });
  }

  const result = [];
  for (const [key, { mat, nome }] of funcSet) {
    const obj = { Matrícula: mat, Nome: nome };
    for (const m of meses) {
      const col = `${m} - Valor`;
      const pgto = (idxPGTO[key] || []).reduce((s, r) => s + (r[col] || 0), 0);
      const desc = (idxDESC[key] || []).reduce((s, r) => s + (r[col] || 0), 0);
      obj[m] = pgto - desc;
    }
    const baseVals = mesesBaseline.map(m => obj[m]).filter(v => v != null && v !== 0);
    obj.MEDIA_BASELINE   = baseVals.length ? mean(baseVals) : null;
    obj.DESVIO_BASELINE  = baseVals.length >= 2 ? std(baseVals) : null;
    obj.MEDIANA_BASELINE = baseVals.length ? median(baseVals) : null;
    obj.MAD_INDIVIDUAL   = baseVals.length >= 2 ? calcMAD(baseVals) : null;
    obj.LIQUIDO_ALVO     = obj[mesAlvo] || 0;
    obj.VAR_ABS  = obj.MEDIA_BASELINE != null ? obj.LIQUIDO_ALVO - obj.MEDIA_BASELINE : null;
    obj.VAR_PCT  = obj.MEDIA_BASELINE ? ((obj.LIQUIDO_ALVO / obj.MEDIA_BASELINE) - 1) * 100 : null;
    obj.Z_SCORE  = (obj.VAR_ABS != null && obj.DESVIO_BASELINE) ? obj.VAR_ABS / obj.DESVIO_BASELINE : null;
    obj.DATA_RESCISAO = desligados[key] || null;
    result.push(obj);
  }

  // IQR global
  const varAbsAll = result.map(r => r.VAR_ABS).filter(v => v != null).sort((a, b) => a - b);
  let q1 = null, q3 = null, iqr = null;
  if (varAbsAll.length > 4) {
    q1 = quantile(varAbsAll, 0.25); q3 = quantile(varAbsAll, 0.75); iqr = q3 - q1;
  }

  const folhaTotalAlvo = result.reduce((s, r) => s + Math.abs(r[mesAlvo] || 0), 0);
  const materialidade  = folhaTotalAlvo * (configParam.materialidade_pct || 0);
  const metodo = configParam.metodo || "sigma";

  for (const r of result) {
    r.STATUS   = _calcStatus(r, metodo, configParam, q1, q3, iqr, materialidade);
    r.AUDITORIA = _calcAuditoria(r, mesAlvo, configParam);
  }
  return { result, metadata: { metodo, q1, q3, iqr, folhaTotalAlvo, materialidade } };
}

function _calcStatus(r, metodo, cfg, q1, q3, iqr, mat) {
  if (r.DATA_RESCISAO) return "DESLIGADO";
  if (r.MEDIA_BASELINE == null || r.MEDIA_BASELINE === 0)
    return r.LIQUIDO_ALVO !== 0 ? "NOVO_FUNC" : "SEM_DADOS";
  if (r.LIQUIDO_ALVO === 0 && r.MEDIA_BASELINE > 0) return "AUSENTE";
  if (r.LIQUIDO_ALVO < -100) return "NEGATIVO";
  if (r.LIQUIDO_ALVO < 500 && r.MEDIA_BASELINE >= 1000) return "ZERO_SUSPEITO";

  if (metodo === "sigma" && r.Z_SCORE != null) {
    if (Math.abs(r.Z_SCORE) >= cfg.extrema) return "EXTREMA";
    if (Math.abs(r.Z_SCORE) >= cfg.alta)    return "ALTA";
  } else if (metodo === "iqr" && r.VAR_ABS != null && iqr != null && iqr > 0) {
    if (r.VAR_ABS < q1 - cfg.extrema * iqr || r.VAR_ABS > q3 + cfg.extrema * iqr) return "EXTREMA_IQR";
    if (r.VAR_ABS < q1 - cfg.alta    * iqr || r.VAR_ABS > q3 + cfg.alta    * iqr) return "ALTA_IQR";
  } else if (metodo === "mad" && r.MAD_INDIVIDUAL > 0 && r.MEDIANA_BASELINE != null) {
    const zMad = 0.6745 * (r.LIQUIDO_ALVO - r.MEDIANA_BASELINE) / r.MAD_INDIVIDUAL;
    if (Math.abs(zMad) >= cfg.extrema) return "EXTREMA";
    if (Math.abs(zMad) >= cfg.alta)    return "ALTA";
  } else if (metodo === "legacy") {
    const vp = r.VAR_PCT != null ? Math.abs(r.VAR_PCT) : 0;
    const va = r.VAR_ABS != null ? Math.abs(r.VAR_ABS) : 0;
    if (vp >= cfg.extrema_pct && va > cfg.extrema_abs) return "EXTREMA";
    if (vp >= cfg.alta_pct    && va > cfg.alta_abs)    return "ALTA";
  }
  if (mat > 0 && r.VAR_ABS != null && Math.abs(r.VAR_ABS) >= mat) return "MATERIAL";
  if (r.Z_SCORE != null && Math.abs(r.Z_SCORE) >= 2) return "Z_2SIGMA";
  return "OK";
}

function _calcAuditoria(r, mesAlvo, cfg) {
  const s = r.STATUS;
  if (s === "DESLIGADO") {
    const d = r.DATA_RESCISAO ? r.DATA_RESCISAO.toLocaleDateString("pt-BR") : "-";
    return `Desligado em ${d}. Líquido ${fmtBRL(r.LIQUIDO_ALVO)} em ${mesAlvo}.`;
  }
  if (s === "AUSENTE")      return `Sem movimento; baseline ${fmtBRL(r.MEDIA_BASELINE)}. Verificar afastamento ou processamento.`;
  if (s === "NEGATIVO")     return `Líquido negativo: ${fmtBRL(r.LIQUIDO_ALVO)}.`;
  if (s === "ZERO_SUSPEITO") return `Líquido próximo de zero (${fmtBRL(r.LIQUIDO_ALVO)}) com baseline ${fmtBRL(r.MEDIA_BASELINE)}.`;
  if (s === "EXTREMA" || s === "EXTREMA_IQR")
    return `Variação extrema: ${fmtPct(r.VAR_PCT)} (${fmtBRL(r.VAR_ABS)}). Z=${r.Z_SCORE != null ? r.Z_SCORE.toFixed(2) : "-"}.`;
  if (s === "ALTA" || s === "ALTA_IQR")
    return `Variação alta: ${fmtPct(r.VAR_PCT)} (${fmtBRL(r.VAR_ABS)}). Z=${r.Z_SCORE != null ? r.Z_SCORE.toFixed(2) : "-"}.`;
  if (s === "MATERIAL")
    return `Impacto material (>=${((cfg.materialidade_pct||0)*100).toFixed(1)}% da folha): ${fmtBRL(r.VAR_ABS)}.`;
  if (s === "Z_2SIGMA")  return `Fora de 2 sigmas (z=${r.Z_SCORE != null ? r.Z_SCORE.toFixed(2) : "-"}).`;
  if (s === "NOVO_FUNC") return `Sem baseline; líquido ${fmtBRL(r.LIQUIDO_ALVO)}.`;
  if (s === "SEM_DADOS") return "Sem movimento no período.";
  return "Sem desvio relevante.";
}

// ---- Verbas zeradas ----
function verbasZeradas(df, classificacao, mesAlvo, mesesBaseline, nMesesMin = 3, valorMin = 500) {
  const sub = df.filter(r => r["Clas."] === classificacao);
  const colAlvo = `${mesAlvo} - Valor`;
  const colsBase = mesesBaseline.map(m => `${m} - Valor`);
  const agg = {};
  for (const r of sub) {
    const key = r["Código"] + "|||" + r["Descrição"];
    if (!agg[key]) agg[key] = { Código: r["Código"], Descrição: r["Descrição"], vals: {}, alvo: 0 };
    for (const c of colsBase) agg[key].vals[c] = (agg[key].vals[c] || 0) + (r[c] || 0);
    agg[key].alvo += (r[colAlvo] || 0);
  }
  const result = [];
  for (const d of Object.values(agg)) {
    if (Math.abs(d.alvo) > 0.01) continue;
    const baseVals = colsBase.map(c => d.vals[c] || 0);
    const nAtivos = baseVals.filter(v => v >= valorMin).length;
    if (nAtivos < nMesesMin) continue;
    const nonZero = baseVals.filter(v => v !== 0);
    result.push({ Código: d.Código, Descrição: d.Descrição, N_MESES_ATIVOS: nAtivos, MEDIA_BASELINE: nonZero.length ? mean(nonZero) : 0, VALOR_ALVO: 0 });
  }
  return result.sort((a, b) => b.MEDIA_BASELINE - a.MEDIA_BASELINE);
}

// ---- Impacto por verba ----
function impactoPorVerba(df, meses, mesAlvo, mesesBaseline, topN = 20) {
  const colAlvo = `${mesAlvo} - Valor`;
  const colsBase = mesesBaseline.map(m => `${m} - Valor`);
  const agg = {};
  for (const r of df) {
    const key = r["Código"] + "|||" + r["Descrição"] + "|||" + r["Clas."];
    if (!agg[key]) agg[key] = { Código: r["Código"], Descrição: r["Descrição"], "Clas.": r["Clas."], vals: {}, alvo: 0 };
    for (const c of colsBase) agg[key].vals[c] = (agg[key].vals[c] || 0) + (r[c] || 0);
    agg[key].alvo += (r[colAlvo] || 0);
  }
  const items = Object.values(agg).map(d => {
    const bv = colsBase.map(c => d.vals[c] || 0).filter(v => v !== 0);
    const mediaBase = bv.length ? mean(bv) : 0;
    const varAbs = d.alvo - mediaBase;
    return { Código: d.Código, Descrição: d.Descrição, "Clas.": d["Clas."], MEDIA_BASELINE: mediaBase,
      VALOR_ALVO: d.alvo, VAR_ABS: varAbs, VAR_PCT: mediaBase ? ((d.alvo / mediaBase) - 1) * 100 : null, IMPACTO_ABS: Math.abs(varAbs) };
  });
  const top = items.sort((a, b) => b.IMPACTO_ABS - a.IMPACTO_ABS).slice(0, topN);
  const ordemClas = { PGTO: 0, DESC: 1, OUTRO: 2 };
  return top.sort((a, b) => {
    const oa = ordemClas[a["Clas."]] ?? 99, ob = ordemClas[b["Clas."]] ?? 99;
    return oa !== ob ? oa - ob : b.IMPACTO_ABS - a.IMPACTO_ABS;
  });
}

// ---- HC e total por verba ----
function hcETotalPorVerba(df, meses, codigosVerba) {
  if (!Array.isArray(codigosVerba)) codigosVerba = [codigosVerba];
  const alvo = new Set(codigosVerba.map(normCod));
  const sub = df.filter(r => alvo.has(normCod(r["Código"])));
  return meses.map(m => {
    const col = `${m} - Valor`;
    const ativos = sub.filter(r => (r[col] || 0) !== 0);
    const mats = new Set(ativos.map(getMatricula));
    const hc = mats.size;
    const total = ativos.reduce((s, r) => s + (r[col] || 0), 0);
    return { Mes: m, HC: hc, Total: total, Media: hc ? total / hc : 0 };
  });
}
