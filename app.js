(() => {
  'use strict';

  const PREVIEW_LIMIT = 100;
  const SOURCE_COLUMN = 'arquivo_origem';
  const DELIMITER_NAMES = { ';': 'ponto e vírgula', ',': 'vírgula', '\t': 'tabulação', '|': 'barra vertical' };

  const els = {
    input: document.getElementById('file-input'),
    dropzone: document.getElementById('dropzone'),
    list: document.getElementById('file-list'),
    empty: document.getElementById('file-empty'),
    clear: document.getElementById('btn-clear'),
    source: document.getElementById('opt-source'),
    outDelim: document.getElementById('opt-delimiter'),
    filename: document.getElementById('opt-filename'),
    download: document.getElementById('btn-download'),
    summary: document.getElementById('summary'),
    preview: document.getElementById('preview'),
    note: document.getElementById('preview-note'),
    status: document.getElementById('status'),
  };

  const state = { files: [], result: null };
  let nextId = 1;
  const fmt = (n) => n.toLocaleString('pt-BR');

  /* ---------- Leitura de arquivos ---------- */

  function decode(buffer) {
    try {
      return { text: new TextDecoder('utf-8', { fatal: true }).decode(buffer), encoding: 'UTF-8' };
    } catch {
      // Arquivos exportados pelo Excel/ERP no Brasil costumam vir em Windows-1252
      return { text: new TextDecoder('windows-1252').decode(buffer), encoding: 'Windows-1252' };
    }
  }

  function detectDelimiter(text) {
    const counts = { ';': 0, ',': 0, '\t': 0, '|': 0 };
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (c === '"') { inQuotes = !inQuotes; continue; }
      if (inQuotes) continue;
      if (c === '\n' || c === '\r') {
        if (Object.values(counts).some((v) => v > 0)) break;
        continue;
      }
      if (c in counts) counts[c]++;
    }
    let best = ';';
    let max = 0;
    for (const [d, n] of Object.entries(counts)) {
      if (n > max) { max = n; best = d; }
    }
    return best;
  }

  function parseCSV(text, delimiter) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    let i = 0;
    const n = text.length;

    const pushRow = () => {
      if (!(row.length === 1 && row[0].trim() === '')) rows.push(row);
      row = [];
    };

    while (i < n) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"' && field === '') { inQuotes = true; i++; continue; }
      if (c === delimiter) { row.push(field); field = ''; i++; continue; }
      if (c === '\r' || c === '\n') {
        row.push(field); field = '';
        pushRow();
        if (c === '\r' && text[i + 1] === '\n') i++;
        i++; continue;
      }
      field += c; i++;
    }
    if (field !== '' || row.length) { row.push(field); pushRow(); }
    return rows;
  }

  async function readFile(file) {
    const buffer = await file.arrayBuffer();
    let { text, encoding } = decode(buffer);
    text = text.replace(/^\uFEFF/, '');
    const delimiter = detectDelimiter(text);
    const rows = parseCSV(text, delimiter);
    if (!rows.length) throw new Error('o arquivo está vazio');
    const header = rows[0].map((h, i) => h.trim() || `coluna_${i + 1}`);
    return { id: nextId++, name: file.name, encoding, delimiter, header, rows: rows.slice(1) };
  }

  async function addFiles(fileList) {
    const files = Array.from(fileList);
    const accepted = files.filter((f) => /\.(csv|txt)$/i.test(f.name));
    const ignored = files.length - accepted.length;
    const errors = [];
    let added = 0;

    for (const file of accepted) {
      try {
        state.files.push(await readFile(file));
        added++;
      } catch (err) {
        errors.push(`${file.name}: ${err.message}`);
      }
    }

    const parts = [];
    if (added) parts.push(`${added} ${added === 1 ? 'arquivo adicionado' : 'arquivos adicionados'}`);
    if (ignored) parts.push(`${ignored} ignorado${ignored > 1 ? 's' : ''} por não ser CSV ou TXT`);
    if (errors.length) parts.push(`não foi possível ler ${errors.join('; ')}`);
    if (parts.length) showStatus(parts.join('. ') + '.');

    update();
  }

  /* ---------- Consolidação ---------- */

  const normalize = (s) => s.trim().toLowerCase();

  function consolidate() {
    const mode = document.querySelector('input[name="mode"]:checked').value;
    const withSource = els.source.checked;
    const info = new Map();
    const rows = [];
    const starts = new Set();
    let columns = [];

    if (mode === 'name') {
      const index = new Map();
      // Monta a união das colunas na ordem em que aparecem; nomes repetidos no
      // mesmo arquivo viram colunas distintas.
      const keysPerFile = state.files.map((f) => {
        const seen = new Map();
        return f.header.map((h) => {
          const base = normalize(h);
          const occ = seen.get(base) || 0;
          seen.set(base, occ + 1);
          const key = `${base}\u0000${occ}`;
          if (!index.has(key)) { index.set(key, columns.length); columns.push(h); }
          return key;
        });
      });

      state.files.forEach((f, fi) => {
        const map = keysPerFile[fi].map((k) => index.get(k));
        const present = new Set(map);
        let extra = 0;
        starts.add(rows.length);
        for (const r of f.rows) {
          const out = new Array(columns.length).fill('');
          const len = Math.min(r.length, map.length);
          for (let i = 0; i < len; i++) out[map[i]] = r[i];
          if (r.length > map.length && r.slice(map.length).some((v) => v !== '')) extra++;
          rows.push(withSource ? [f.name, ...out] : out);
        }
        info.set(f.id, {
          missing: columns.filter((_, i) => !present.has(i)),
          extra,
        });
      });
    } else {
      let width = 0;
      for (const f of state.files) {
        width = Math.max(width, f.header.length);
        for (const r of f.rows) if (r.length > width) width = r.length;
      }
      const first = state.files[0] ? state.files[0].header : [];
      columns = Array.from({ length: width }, (_, i) => first[i] || `coluna_${i + 1}`);
      const firstKey = first.map(normalize).join('|');

      state.files.forEach((f, fi) => {
        starts.add(rows.length);
        for (const r of f.rows) {
          const out = new Array(width).fill('');
          for (let i = 0; i < r.length; i++) out[i] = r[i];
          rows.push(withSource ? [f.name, ...out] : out);
        }
        info.set(f.id, {
          headerDiffers: fi > 0 && f.header.map(normalize).join('|') !== firstKey,
        });
      });
    }

    if (withSource) columns = [SOURCE_COLUMN, ...columns];
    return { columns, rows, info, starts, withSource };
  }

  function toCSV({ columns, rows }, delimiter) {
    const needsQuote = new RegExp(`["\\r\\n${delimiter === '\t' ? '\\t' : '\\' + delimiter}]|^\\s|\\s$`);
    const esc = (v) => {
      const s = v == null ? '' : String(v);
      return needsQuote.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const lines = new Array(rows.length + 1);
    lines[0] = columns.map(esc).join(delimiter);
    for (let i = 0; i < rows.length; i++) lines[i + 1] = rows[i].map(esc).join(delimiter);
    return lines.join('\r\n');
  }

  function download() {
    if (!state.result || !state.result.rows.length) return;
    const delimiter = els.outDelim.value === '\\t' ? '\t' : els.outDelim.value;
    const csv = toCSV(state.result, delimiter);
    // BOM para o Excel abrir acentos corretamente
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
    let name = (els.filename.value || 'consolidado').trim().replace(/[\\/:*?"<>|]/g, '_');
    if (!/\.(csv|txt)$/i.test(name)) name += '.csv';

    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    showStatus(`${name} baixado com ${fmt(state.result.rows.length)} linhas.`);
  }

  /* ---------- Interface ---------- */

  function el(tag, props = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(props)) {
      if (k === 'class') node.className = v;
      else if (k === 'style') node.setAttribute('style', v);
      else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
      else if (v !== false && v != null) node.setAttribute(k, v === true ? '' : v);
    }
    for (const c of children.flat()) {
      if (c != null && c !== false) node.append(c);
    }
    return node;
  }

  function move(id, dir) {
    const i = state.files.findIndex((f) => f.id === id);
    const j = i + dir;
    if (i < 0 || j < 0 || j >= state.files.length) return;
    [state.files[i], state.files[j]] = [state.files[j], state.files[i]];
    update();
  }

  function remove(id) {
    state.files = state.files.filter((f) => f.id !== id);
    update();
  }

  function renderList() {
    const total = state.files.reduce((s, f) => s + f.rows.length, 0) || 1;
    els.list.replaceChildren(...state.files.map((f, i) => {
      const inf = state.result.info.get(f.id) || {};
      const warnings = [];
      if (inf.missing && inf.missing.length) {
        const list = inf.missing.slice(0, 4).join(', ') + (inf.missing.length > 4 ? ` e mais ${inf.missing.length - 4}` : '');
        warnings.push(`Sem ${inf.missing.length === 1 ? 'a coluna' : 'as colunas'} ${list}; esses campos ficam vazios.`);
      }
      if (inf.extra) warnings.push(`${fmt(inf.extra)} linha${inf.extra > 1 ? 's têm' : ' tem'} mais valores que o cabeçalho; o excedente foi descartado.`);
      if (inf.headerDiffers) warnings.push('O cabeçalho é diferente do primeiro arquivo. Confira se as colunas estão na mesma ordem.');

      const pct = (f.rows.length / total) * 100;
      return el('li', { class: 'file' },
        el('span', { class: 'file-name' }, f.name),
        el('div', { class: 'file-actions' },
          el('button', { class: 'icon-btn', type: 'button', 'aria-label': `Mover ${f.name} para cima`, title: 'Mover para cima', disabled: i === 0, onclick: () => move(f.id, -1) }, '↑'),
          el('button', { class: 'icon-btn', type: 'button', 'aria-label': `Mover ${f.name} para baixo`, title: 'Mover para baixo', disabled: i === state.files.length - 1, onclick: () => move(f.id, 1) }, '↓'),
          el('button', { class: 'icon-btn remove', type: 'button', 'aria-label': `Remover ${f.name}`, title: 'Remover', onclick: () => remove(f.id) }, '×'),
        ),
        el('span', { class: 'file-meta' },
          `${fmt(f.rows.length)} linhas, ${f.header.length} colunas, separador ${DELIMITER_NAMES[f.delimiter]}, ${f.encoding}`),
        el('div', { class: 'file-bar', title: `${pct.toFixed(1).replace('.', ',')}% das linhas` },
          el('span', { style: `width:${Math.max(pct, 0.5)}%` })),
        warnings.map((w) => el('p', { class: 'file-warn' }, w)),
      );
    }));
  }

  function renderPreview() {
    const { columns, rows, starts, withSource } = state.result;

    if (!state.files.length) {
      els.preview.replaceChildren(el('tbody', {}, el('tr', { class: 'table-placeholder' },
        el('td', {}, 'Adicione arquivos para ver aqui como a base consolidada vai ficar.'))));
      els.summary.replaceChildren();
      els.note.textContent = '';
      return;
    }

    els.summary.replaceChildren(
      el('b', {}, fmt(rows.length)), ' linhas, ',
      el('b', {}, fmt(columns.length)), ' colunas, ',
      el('b', {}, fmt(state.files.length)), state.files.length === 1 ? ' arquivo' : ' arquivos',
    );

    const shown = rows.slice(0, PREVIEW_LIMIT);
    els.preview.replaceChildren(
      el('thead', {}, el('tr', {}, columns.map((c) => el('th', { scope: 'col', title: c }, c)))),
      el('tbody', {}, shown.map((r, ri) =>
        el('tr', { class: ri > 0 && starts.has(ri) ? 'boundary' : null },
          r.map((v, ci) => el('td', {
            class: withSource && ci === 0 ? 'src' : (v === '' ? 'blank' : null),
            title: v || null,
          }, v))))),
    );

    els.note.textContent = rows.length > PREVIEW_LIMIT
      ? `Mostrando as primeiras ${PREVIEW_LIMIT} de ${fmt(rows.length)} linhas. O arquivo baixado traz todas. A linha verde marca o início de cada arquivo.`
      : rows.length ? 'A linha verde marca o início de cada arquivo.' : 'Os arquivos têm só cabeçalho, sem linhas de dados.';
  }

  function update() {
    state.result = consolidate();
    const has = state.files.length > 0;
    els.empty.hidden = has;
    els.clear.disabled = !has;
    els.download.disabled = !has || !state.result.rows.length;
    renderList();
    renderPreview();
  }

  let statusTimer;
  function showStatus(msg) {
    els.status.textContent = msg;
    els.status.classList.add('is-visible');
    clearTimeout(statusTimer);
    statusTimer = setTimeout(() => els.status.classList.remove('is-visible'), 4500);
  }

  /* ---------- Eventos ---------- */

  els.input.addEventListener('change', (e) => {
    addFiles(e.target.files);
    e.target.value = '';
  });

  ['dragenter', 'dragover'].forEach((type) =>
    els.dropzone.addEventListener(type, (e) => {
      e.preventDefault();
      els.dropzone.classList.add('is-over');
    }));
  ['dragleave', 'drop'].forEach((type) =>
    els.dropzone.addEventListener(type, () => els.dropzone.classList.remove('is-over')));
  els.dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });
  // Evita que o navegador abra o arquivo se ele for solto fora da área
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());

  document.querySelectorAll('input[name="mode"]').forEach((r) => r.addEventListener('change', update));
  els.source.addEventListener('change', update);
  els.clear.addEventListener('click', () => { state.files = []; update(); showStatus('Lista limpa.'); });
  els.download.addEventListener('click', download);

  update();
})();
