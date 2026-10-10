/* Seleção numérica reutilizável — sem dependências e sem mutação dos dados. */
(function (root) {
  'use strict';
  // Conversão decimal por dígitos: inclusive notação científica; arredondamento
  // financeiro na terceira casa, sem multiplicação binária por 100.
  function cents(value) {
    if (value == null || value === '') return 0n;
    if (typeof value === 'bigint') return value;
    const s = String(value).trim();
    const m = /^([+-]?)(\d+)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(s);
    if (!m) throw new TypeError('Valor monetário deve estar normalizado: ' + s);
    const digits = (m[2] + (m[3] || '')).replace(/^0+(?=\d)/, '');
    const shift = 2 + Number(m[4] || 0) - (m[3] || '').length;
    let n;
    if (shift >= 0) n = BigInt(digits) * 10n ** BigInt(shift);
    else {
      const divisor = 10n ** BigInt(-shift), raw = BigInt(digits);
      n = raw / divisor + (raw % divisor * 2n >= divisor ? 1n : 0n);
    }
    return m[1] === '-' ? -n : n;
  }
  function formatCents(n) {
    const neg = n < 0n; if (neg) n = -n;
    const whole = (n / 100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (neg ? '−' : '') + 'R$ ' + whole + ',' + (n % 100n).toString().padStart(2, '0');
  }
  function payroll(value, classification, code, description, isLiquid) {
    const k = String(classification || '').trim().toUpperCase();
    let n = cents(value);
    if (isLiquid || String(code).replace(/^0+/, '') === '9950')
      return { 'Líquido informado (referência)': n };
    if (/DESC/.test(k)) { if (n > 0n) n = -n; return { 'Folha (proventos − descontos)': n }; }
    if (/PGTO|PGT|PAGAM/.test(k)) return { 'Folha (proventos − descontos)': n };
    // BASE/PROV/ENCAR/OUTRO não são remuneração e não são somadas entre rubricas
    // distintas. Preserva o sinal da fonte e identifica cada natureza.
    return { [k + ' · ' + code + ' — ' + (description || 'Sem descrição')]: n };
  }
  function combine(parts) {
    const out = {};
    for (const part of parts) for (const [key, n] of Object.entries(part)) out[key] = (out[key] || 0n) + n;
    return out;
  }
  let sourceIds = new WeakMap(), sequence = 0, dataset = 0;
  function sourceId(record) {
    if (!sourceIds.has(record)) sourceIds.set(record, 'base:' + dataset + ':' + (++sequence));
    return sourceIds.get(record);
  }
  class SelectionModel {
    constructor() { this.selected = new Set(); this.records = new Map(); }
    setRecords(records, key, values) {
      const next = new Map();
      for (const record of records) {
        const id = String(key(record));
        if (next.has(id)) throw new Error('Identificador de seleção duplicado: ' + id);
        next.set(id, { record, values: values(record) });
      }
      let removed = 0;
      for (const id of this.selected) if (!next.has(id)) { this.selected.delete(id); removed++; }
      this.records = next;
      return removed;
    }
    set(id, on) { id = String(id); if (!this.records.has(id)) return; if (on) this.selected.add(id); else this.selected.delete(id); }
    toggle(id) { this.set(id, !this.selected.has(String(id))); }
    clear() { this.selected.clear(); }
    summary() {
      const totals = {};
      for (const id of this.selected) {
        const entry = this.records.get(id);
        if (!entry) continue;
        for (const [group, value] of Object.entries(entry.values)) totals[group] = (totals[group] || 0n) + value;
      }
      return { count: this.selected.size, totals };
    }
  }
  const registry = new Map();
  class TableSelection {
    constructor(name) {
      this.name = name; this.model = new SelectionModel(); this.pageIndex = 0;
      this.pageSize = 100; this.pending = new Map(); this.message = ''; this.rows = new Map();
    }
    cancelPending() { for (const p of this.pending.values()) clearTimeout(p.timer); this.pending.clear(); }
    prepare(records, options) {
      this.cancelPending(); this.options = options;
      const old = this.model.records;
      const removed = this.model.setRecords(records, options.key, options.values);
      if (removed) this.message = removed + ' linha(s) removida(s) da seleção por sair do resultado filtrado.';
      const changed = old.size !== this.model.records.size || [...old.keys()].some(k => !this.model.records.has(k));
      if (changed) this.pageIndex = 0;
      this.records = records;
      this.pageIndex = Math.min(this.pageIndex, Math.max(0, Math.ceil(records.length / this.pageSize) - 1));
      this.rows = new Map();
      return records.slice(this.pageIndex * this.pageSize, (this.pageIndex + 1) * this.pageSize);
    }
    bind(tr, record) {
      const id = String(this.options.key(record));
      this.rows.set(tr, id);
      tr.dataset.sumId = id; tr.tabIndex = 0;
      tr.setAttribute('aria-label', 'Selecionar para soma. Espaço alterna; Enter investiga.');
      tr.classList.add('sum-selectable');
      const td = document.createElement('td'); td.className = 'sum-cell';
      const checkbox = document.createElement('input'); checkbox.type = 'checkbox'; checkbox.className = 'sum-check';
      checkbox.setAttribute('aria-label', 'Selecionar linha para soma');
      td.appendChild(checkbox); tr.insertBefore(td, tr.firstChild);
      // Checkbox tem uma única via: change. Não aciona o clique da linha.
      checkbox.addEventListener('click', e => e.stopPropagation());
      checkbox.addEventListener('dblclick', e => e.stopPropagation());
      checkbox.addEventListener('change', () => { this.model.set(id, checkbox.checked); this.message = ''; this.refresh(); });
      tr.addEventListener('click', e => {
        if (e.target.closest('input,button,select,textarea,a,[contenteditable]')) return;
        const pending = this.pending.get(id);
        if (e.detail > 1) { if (pending) clearTimeout(pending.timer); return; }
        if (pending) clearTimeout(pending.timer);
        const before = this.model.selected.has(id);
        const gesture = { before, committed: false };
        // Browser emite click/click/dblclick. Aguarda 360 ms para confirmar
        // clique simples; checkbox e teclado respondem imediatamente.
        gesture.timer = setTimeout(() => {
          if (!this.model.records.has(id)) return;
          this.model.toggle(id); gesture.committed = true; this.message = ''; this.refresh();
        }, 360);
        this.pending.set(id, gesture);
      });
      tr.addEventListener('dblclick', e => {
        if (e.target.closest('input,button,select,textarea,a,[contenteditable]')) return;
        const gesture = this.pending.get(id);
        if (gesture) {
          clearTimeout(gesture.timer);
          // Também atende double-click lento reconhecido pelo sistema.
          if (gesture.committed) { this.model.set(id, gesture.before); this.refresh(); }
          this.pending.delete(id);
        }
        if (this.options.investigate) this.options.investigate(record, tr);
      });
      tr.addEventListener('keydown', e => {
        if (e.target !== tr || e.repeat) return;
        if (e.key === ' ' || e.key === 'Spacebar') {
          e.preventDefault(); this.cancelPending(); this.model.toggle(id); this.message = ''; this.refresh();
        } else if (e.key === 'Enter' && this.options.investigate) {
          e.preventDefault(); this.cancelPending(); this.options.investigate(record, tr);
        }
      });
    }
    finish(table) {
      if (!table) return;
      table.classList.add('sum-table');
      const header = table.querySelector('thead tr');
      if (header && !header.querySelector('.sum-head')) {
        const th = document.createElement('th'); th.className = 'sum-head'; th.textContent = 'Soma'; th.scope = 'col';
        header.insertBefore(th, header.firstChild);
      }
      // Mantém o placeholder cobrindo todas as colunas.
      const placeholder = table.querySelector('tbody td[colspan]');
      if (placeholder) placeholder.colSpan = header ? header.children.length : placeholder.colSpan;
      const wrap = table.closest('.table-wrap,.table-scroll') || table;
      if (!this.bar || !this.bar.isConnected) {
        this.bar = document.createElement('div'); this.bar.className = 'sum-toolbar'; this.bar.dataset.selection = this.name;
      }
      wrap.parentNode.insertBefore(this.bar, wrap);
      this.bar.replaceChildren();
      const label = document.createElement('span'); label.className = 'sum-label'; label.textContent = this.options.label || 'Valores';
      const output = document.createElement('div'); output.className = 'sum-output'; output.setAttribute('aria-live', 'polite'); output.setAttribute('aria-atomic', 'true');
      const clear = document.createElement('button'); clear.type = 'button'; clear.textContent = 'Limpar seleção'; clear.className = 'sum-clear';
      clear.addEventListener('click', () => { this.cancelPending(); this.model.clear(); this.message = 'Seleção limpa.'; this.refresh(); });
      const note = document.createElement('div'); note.className = 'sum-notice'; note.setAttribute('role', 'status');
      const help = document.createElement('small'); help.textContent = 'Clique seleciona · duplo clique investiga · Espaço seleciona · Enter investiga';
      this.bar.append(label, output, clear, note, help);
      if (!this.pager || !this.pager.isConnected) { this.pager = document.createElement('div'); this.pager.className = 'sum-pager'; this.pager.dataset.selectionPager = this.name; }
      wrap.after(this.pager); this.pager.replaceChildren();
      const pages = Math.max(1, Math.ceil(this.records.length / this.pageSize));
      const prev = document.createElement('button'); prev.type = 'button'; prev.textContent = 'Anterior'; prev.disabled = this.pageIndex === 0;
      const next = document.createElement('button'); next.type = 'button'; next.textContent = 'Próxima'; next.disabled = this.pageIndex >= pages - 1;
      const text = document.createElement('span'); text.textContent = `Página ${this.pageIndex + 1} de ${pages} · ${this.records.length} linha(s) filtrada(s)`;
      const size = document.createElement('select'); size.setAttribute('aria-label', 'Linhas por página');
      for (const n of [50, 100, 250, 500]) { const op = document.createElement('option'); op.value = n; op.textContent = n + ' por página'; op.selected = n === this.pageSize; size.appendChild(op); }
      const go = delta => { this.cancelPending(); this.pageIndex += delta; this.options.render(); this.pager?.querySelector(delta < 0 ? 'button' : 'button:nth-of-type(2)')?.focus(); };
      prev.onclick = () => go(-1); next.onclick = () => go(1);
      size.onchange = () => { this.pageSize = Number(size.value); this.pageIndex = 0; this.options.render(); };
      this.pager.append(prev, text, next, size);
      this.refresh();
    }
    refresh() {
      for (const [tr, id] of this.rows) {
        const on = this.model.selected.has(id);
        tr.classList.toggle('sum-selected', on); tr.setAttribute('aria-selected', String(on));
        const chk = tr.querySelector('.sum-check'); if (chk) chk.checked = on;
      }
      if (!this.bar) return;
      const summary = this.model.summary(); const output = this.bar.querySelector('.sum-output');
      output.replaceChildren();
      const count = document.createElement('strong'); count.className = 'sum-count'; count.textContent = 'Linhas selecionadas: ' + summary.count; output.appendChild(count);
      const totals = Object.entries(summary.totals);
      if (!totals.length) totals.push([this.options.emptyGroup || 'Valores', 0n]);
      for (const [group, value] of totals) {
        const span = document.createElement('span'); span.className = 'sum-total'; span.dataset.group = group; span.dataset.cents = String(value);
        span.textContent = 'Total selecionado · ' + group + ': ' + formatCents(value); output.appendChild(span);
      }
      this.bar.querySelector('.sum-notice').textContent = this.message;
      this.bar.querySelector('.sum-clear').disabled = !summary.count && !this.pending.size;
    }
    reset() { this.cancelPending(); this.model.clear(); this.model.records.clear(); this.pageIndex = 0; this.message = ''; if (this.bar) this.refresh(); }
  }
  function table(name) { if (!registry.has(name)) registry.set(name, new TableSelection(name)); return registry.get(name); }
  function newDataset() { dataset++; sequence = 0; sourceIds = new WeakMap(); for (const view of registry.values()) view.reset(); }
  function openDetail(title, record) {
    const dialog = document.createElement('dialog'); dialog.className = 'sum-detail';
    const h = document.createElement('h3'); h.textContent = title;
    const dl = document.createElement('dl');
    for (const [key, value] of Object.entries(record)) {
      const dt = document.createElement('dt'); dt.textContent = key;
      const dd = document.createElement('dd');
      dd.textContent = value && typeof value === 'object' ? JSON.stringify(value, (_, v) => typeof v === 'bigint' ? formatCents(v) : v, 2) : String(value ?? '—');
      dl.append(dt, dd);
    }
    const close = document.createElement('button'); close.textContent = 'Fechar'; close.type = 'button'; close.onclick = () => dialog.close();
    dialog.append(h, dl, close); dialog.addEventListener('close', () => dialog.remove());
    document.body.appendChild(dialog); dialog.showModal(); close.focus();
  }
  root.RowSelection = { cents, formatCents, payroll, combine, sourceId, SelectionModel, table, newDataset, openDetail };
  if (typeof module !== 'undefined') module.exports = root.RowSelection;
})(typeof window !== 'undefined' ? window : globalThis);
