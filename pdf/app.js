/*
 * Cofre PDF: anonimização local de documentos PDF.
 * Ana Marta Soares, Especialista em IA aplicada ao Direito.
 * O PDF é lido e processado neste computador; não há envio de conteúdo.
 */
import * as pdfjsLib from './lib/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./lib/pdf.worker.min.mjs', import.meta.url).href;

const E = window.AnonEngine, D = window.DocxUtil, L = window.Licenca;
const $ = (id) => document.getElementById(id);
const LIC_KEY = 'cofre.pdf.licenca';
const VIEW_SCALE_MAX = 1.6;
const EXPORT_SCALE = 2.5;

const st = {
  licensed: false, pdf: null, name: '', pages: [], lines: [], entities: [], manual: {}, flashT: 0
};

/* ------------------------------------------------------------- licença */
function storeGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
function storeSet(k, v) { try { v ? localStorage.setItem(k, v) : localStorage.removeItem(k); return true; } catch (e) { return false; } }
async function revoked() {
  try { const r = await fetch('../revogadas.json?t=' + Date.now(), { cache: 'no-store' }); return r.ok ? ((await r.json()).revogadas || []) : []; }
  catch (e) { return []; }
}
async function checkCode(code) {
  const v = L.verify(code);
  if (!v.ok) return v;
  if (!L.allows('pdf', v.tipo)) return { ok: false, motivo: 'Este código é de outro produto (Cofre para Word) e não activa o Cofre PDF.' };
  if ((await revoked()).includes(v.num)) return { ok: false, motivo: 'Esta licença foi desactivada. Contacte Ana Marta Soares.' };
  return v;
}
function setLicensed(v) {
  st.licensed = !!(v && v.ok);
  $('licGate').hidden = st.licensed;
  $('work').hidden = !st.licensed;
  $('licInfo').textContent = st.licensed ? `${v.tipoNome} n.º ${v.num}, ${v.validade ? 'válida até ' + v.validade.toLocaleDateString('pt-PT') : 'sem prazo de validade'}.` : '';
  if (st.licensed && st.pdf && !st.entities.length) analyze();
}
async function initLicense() {
  const c = storeGet(LIC_KEY);
  if (!c) return setLicensed(null);
  const v = await checkCode(c);
  if (!v.ok) { storeSet(LIC_KEY, null); status('licStatus', v.motivo, 'err'); }
  setLicensed(v);
}
async function activate() {
  const code = $('licCode').value.replace(/\s+/g, '');
  if (!code) return status('licStatus', 'Cole o código de licença.', 'err');
  status('licStatus', 'A verificar…', 'busy');
  const v = await checkCode(code);
  if (!v.ok) return status('licStatus', v.motivo, 'err');
  storeSet(LIC_KEY, code); $('licCode').value = ''; status('licStatus', '');
  setLicensed(v);
}

/* ---------------------------------------------------------- utilitários */
function status(id, msg, kind) { const el = $(id); el.className = 'status' + (kind ? ' ' + kind : ''); el.textContent = msg || ''; }
function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function plural(n, a, b) { return n + ' ' + (n === 1 ? a : b); }
function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
function baseName() { return (st.name || 'documento.pdf').replace(/\.pdf$/i, ''); }

/* ------------------------------------------------------------- abertura */
async function openFile(file) {
  if (!file) return;
  if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') { alert('Seleccione um ficheiro PDF.'); return; }
  st.name = file.name;
  $('fileName').textContent = file.name;
  const buf = await file.arrayBuffer();
  try {
    st.pdf = await pdfjsLib.getDocument({ data: buf, standardFontDataUrl: new URL('./lib/standard_fonts/', import.meta.url).href, isEvalSupported: false }).promise;
  } catch (e) {
    const msg = /password/i.test(e && e.name || '') ? 'Este PDF está protegido por palavra-passe. Retire a protecção antes de o anonimizar.' : 'Não foi possível abrir este PDF.';
    alert(msg); return;
  }
  st.pages = []; st.lines = []; st.entities = []; st.manual = {};
  $('empty').hidden = true; $('pages').hidden = false; $('pages').innerHTML = '';
  await extract();
  await renderPages();
  if (st.licensed) analyze();
  else $('summary').textContent = 'Active a licença para identificar os dados.';
}

const measureCtx = document.createElement('canvas').getContext('2d');
function charOffsets(str, family) {
  measureCtx.font = '100px ' + family;
  const total = measureCtx.measureText(str).width || str.length;
  const out = [0];
  for (let i = 1; i <= str.length; i++) out.push(i === str.length ? 1 : measureCtx.measureText(str.slice(0, i)).width / total);
  return out;
}

/* Extrai o texto de cada página, agrupado em linhas, com a posição de cada carácter. */
async function extract() {
  for (let n = 1; n <= st.pdf.numPages; n++) {
    const page = await st.pdf.getPage(n);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    const pinfo = { n, page, w: vp.width, h: vp.height, lines: [] };
    let cur = null, prev = null;
    const rawLines = [];
    const flush = () => { if (cur && cur.text.trim()) rawLines.push(cur); cur = null; };
    for (const it of tc.items) {
      if (typeof it.str !== 'string') continue;
      const tx = pdfjsLib.Util.transform(vp.transform, it.transform);
      const fh = Math.hypot(tx[2], tx[3]) || 10;
      const x = tx[4], yb = tx[5], w = it.width;
      if (it.str.length) {
        if (!cur || !prev || Math.abs(yb - prev.yb) > fh * 0.5 || x < prev.xe - fh) { flush(); cur = { page: n, text: '', map: [] }; }
        else {
          const gap = x - prev.xe;
          if (gap > fh * 0.12 && !/\s$/.test(cur.text) && !/^\s/.test(it.str)) { cur.text += ' '; cur.map.push(null); }
        }
        const len = it.str.length;
        // posição de cada carácter, estimada pela largura relativa dos glifos na família de letra do PDF
        const fam = (tc.styles[it.fontName] && tc.styles[it.fontName].fontFamily) || 'serif';
        const cum = charOffsets(it.str, fam);
        for (let i = 0; i < len; i++) {
          cur.text += it.str[i];
          cur.map.push({ x0: x + w * cum[i], x1: x + w * cum[i + 1], yt: yb - fh * 0.88, yb2: yb + fh * 0.24, pad: fh * 0.12, bl: yb, fh });
        }
        prev = { yb, xe: x + w };
      }
      if (it.hasEOL) { flush(); prev = null; }
    }
    flush();
    // junta linhas consecutivas do mesmo parágrafo, para detectar dados partidos entre linhas
    let block = null, last = null;
    for (const ln of rawLines) {
      const first = ln.map.find(Boolean);
      const h = first ? (first.yb2 - first.yt) : 12;
      const y = first ? first.yt : 0;
      if (block && last && y - last.y > 0 && y - last.y <= h * 1.55) {
        block.text += ' '; block.map.push(null);
        block.text += ln.text; block.map.push(...ln.map);
      } else {
        if (block) { pinfo.lines.push(block); st.lines.push(block); }
        block = { page: n, text: ln.text, map: ln.map.slice() };
      }
      last = { y };
    }
    if (block) { pinfo.lines.push(block); st.lines.push(block); }
    st.pages.push(pinfo);
  }
}

/* --------------------------------------------------------- visualização */
async function renderPages() {
  const host = $('pages');
  const avail = Math.max(320, $('viewer').clientWidth - 60);
  for (const p of st.pages) {
    const scale = Math.min(VIEW_SCALE_MAX, avail / p.w);
    const dpr = window.devicePixelRatio || 1;
    const vp = p.page.getViewport({ scale: scale * dpr });
    const wrap = document.createElement('div');
    wrap.className = 'page';
    wrap.style.width = (p.w * scale) + 'px'; wrap.style.height = (p.h * scale) + 'px';
    const cv = document.createElement('canvas');
    cv.width = vp.width; cv.height = vp.height;
    cv.style.width = (p.w * scale) + 'px'; cv.style.height = (p.h * scale) + 'px';
    const num = document.createElement('div'); num.className = 'num'; num.textContent = 'Página ' + p.n;
    const ov = document.createElement('div'); ov.className = 'ov';
    wrap.append(num, cv, ov); host.appendChild(wrap);
    p.scale = scale; p.ov = ov; p.wrap = wrap;
    await p.page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
    bindDraw(p);
  }
  const noText = st.lines.length === 0;
  if (noText) $('summary').textContent = 'Este PDF não tem texto pesquisável (provavelmente é digitalizado). Arraste o rato sobre as zonas a rasurar.';
}

/* Caixas manuais: arrastar para criar, clicar para retirar. */
function bindDraw(p) {
  let start = null, draft = null;
  p.ov.addEventListener('mousedown', (ev) => {
    if (ev.target.classList.contains('man')) return;
    const r = p.ov.getBoundingClientRect();
    start = { x: ev.clientX - r.left, y: ev.clientY - r.top };
    draft = document.createElement('div'); draft.className = 'box draft'; p.ov.appendChild(draft);
  });
  window.addEventListener('mousemove', (ev) => {
    if (!start) return;
    const r = p.ov.getBoundingClientRect();
    const x = ev.clientX - r.left, y = ev.clientY - r.top;
    Object.assign(draft.style, { left: Math.min(x, start.x) + 'px', top: Math.min(y, start.y) + 'px', width: Math.abs(x - start.x) + 'px', height: Math.abs(y - start.y) + 'px' });
  });
  window.addEventListener('mouseup', (ev) => {
    if (!start) return;
    const r = p.ov.getBoundingClientRect();
    const x = ev.clientX - r.left, y = ev.clientY - r.top;
    const bx = Math.min(x, start.x) / p.scale, by = Math.min(y, start.y) / p.scale, bw = Math.abs(x - start.x) / p.scale, bh = Math.abs(y - start.y) / p.scale;
    draft.remove(); start = null; draft = null;
    if (bw < 4 || bh < 4) return;
    (st.manual[p.n] = st.manual[p.n] || []).push({ x: bx, y: by, w: bw, h: bh });
    drawBoxes();
  });
}

/* ------------------------------------------------------------- análise */
function analyze() {
  if (!st.lines.length) { drawBoxes(); updateButtons(); return; }
  const res = E.analyze(st.lines.map((l) => l.text), {});
  st.entities = res.entities;
  computeBoxes(); renderList(); drawBoxes(); updateButtons();
}
function computeBoxes() {
  // as formas mais longas são localizadas primeiro; as mais curtas não podem sobrepor-se a elas
  const units = [];
  for (const en of st.entities) { en.boxes = []; en.occ = 0; en.ranges = []; for (const u of E.unitsOf(en)) units.push({ en, u, needle: u.prefix + u.text }); }
  units.sort((a, b) => b.needle.length - a.needle.length);
  const covered = new Map();
  const free = (ln, a, b) => (covered.get(ln) || []).every(([x, y]) => b <= x || a >= y);
  const take = (ln, a, b) => { if (!covered.has(ln)) covered.set(ln, []); covered.get(ln).push([a, b]); };
  for (const { en, u, needle } of units) {
    {
      const ww = E.isWordy(needle);
      st.lines.forEach((ln) => {
        for (const i of E.wordSearch(ln.text, needle, !en.manual, ww)) {
          if (!free(ln, i + u.prefix.length, i + needle.length)) continue;
          take(ln, i + u.prefix.length, i + needle.length);
          const a = i + u.prefix.length, b = i + needle.length; en.occ++; en.ranges.push({ ln, a, b });
          let box = null;
          for (let k = a; k < b; k++) {
            const m = ln.map[k];
            if (!m) continue;
            // nova caixa quando o dado continua na linha seguinte
            if (box && (Math.abs(m.yt - box.yt) > (box.yb - box.yt) * 0.5 || m.x0 < box.x0 - 1)) { en.boxes.push(box); box = null; }
            if (!box) box = { page: ln.page, x0: m.x0 - m.pad, x1: m.x1 + m.pad, yt: m.yt, yb: m.yb2 };
            else { box.x0 = Math.min(box.x0, m.x0 - m.pad); box.x1 = Math.max(box.x1, m.x1 + m.pad); box.yt = Math.min(box.yt, m.yt); box.yb = Math.max(box.yb, m.yb2); }
          }
          if (box) en.boxes.push(box);
        }
      });
    }
    en.count = en.boxes.length || en.count;
  }
}
function renderList() {
  const list = $('entityList'); list.innerHTML = '';
  const ents = st.entities;
  const total = ents.reduce((a, e) => a + (e.occ || 0), 0);
  $('summary').textContent = ents.length
    ? `${plural(ents.length, 'elemento distinto', 'elementos distintos')} em ${plural(total, 'ocorrência', 'ocorrências')}. Desmarque o que não deve ser rasurado.`
    : 'Não foram identificados dados pessoais. Pode acrescentar termos ou rasurar zonas manualmente.';
  $('listTools').hidden = !ents.length;
  const groups = {};
  ents.forEach((en) => (groups[en.type] = groups[en.type] || []).push(en));
  for (const t of E.TYPE_ORDER) {
    const g = groups[t]; if (!g) continue;
    const wrap = document.createElement('div'); wrap.className = 'group';
    const head = document.createElement('div'); head.className = 'group-h';
    head.innerHTML = `<input type="checkbox" ${g.every((e) => e.selected) ? 'checked' : ''}><span>${esc(E.TYPES[t].name)}</span><span class="cnt">${g.length}</span>`;
    const gc = head.querySelector('input');
    gc.addEventListener('change', () => { g.forEach((e) => (e.selected = gc.checked)); renderList(); drawBoxes(); });
    wrap.appendChild(head);
    for (const en of g) {
      const row = document.createElement('div'); row.className = 'item' + (en.selected ? '' : ' off');
      row.innerHTML = `<input type="checkbox" ${en.selected ? 'checked' : ''}><div class="txt"><div class="v">${esc(en.canonical)}</div><span class="chip">${esc(en.label)}</span><span class="times">${plural(en.occ, 'ocorrência', 'ocorrências')}</span></div><button class="loc">Ver</button>`;
      row.querySelector('input').addEventListener('change', (ev) => { en.selected = ev.target.checked; row.classList.toggle('off', !en.selected); gc.checked = g.every((e) => e.selected); drawBoxes(); });
      row.querySelector('.loc').addEventListener('click', () => locate(en));
      wrap.appendChild(row);
    }
    list.appendChild(wrap);
  }
}
function drawBoxes() {
  for (const p of st.pages) {
    p.ov.querySelectorAll('.box').forEach((b) => b.remove());
    for (const en of st.entities) {
      if (!en.selected) continue;
      for (const b of en.boxes || []) {
        if (b.page !== p.n) continue;
        const d = document.createElement('div'); d.className = 'box auto'; d.dataset.e = en.id; d.title = en.label;
        Object.assign(d.style, { left: (b.x0 - 1) * p.scale + 'px', top: (b.yt - 1) * p.scale + 'px', width: (b.x1 - b.x0 + 2) * p.scale + 'px', height: (b.yb - b.yt + 2) * p.scale + 'px' });
        p.ov.appendChild(d);
      }
    }
    (st.manual[p.n] || []).forEach((m, i) => {
      const d = document.createElement('div'); d.className = 'box man'; d.title = 'Caixa manual: clique para retirar';
      Object.assign(d.style, { left: m.x * p.scale + 'px', top: m.y * p.scale + 'px', width: m.w * p.scale + 'px', height: m.h * p.scale + 'px' });
      d.addEventListener('click', (ev) => { ev.stopPropagation(); st.manual[p.n].splice(i, 1); drawBoxes(); });
      p.ov.appendChild(d);
    });
  }
}
function locate(en) {
  const b = (en.boxes || [])[0]; if (!b) return;
  const p = st.pages[b.page - 1];
  p.wrap.scrollIntoView({ behavior: 'smooth', block: 'start' });
  const els = [...p.ov.querySelectorAll(`.box.auto[data-e="${en.id}"]`)];
  els.forEach((e) => e.classList.add('flash'));
  clearTimeout(st.flashT); st.flashT = setTimeout(() => els.forEach((e) => e.classList.remove('flash')), 1800);
}
function addTerm() {
  const t = $('manualTerm').value.trim(); if (!t || !st.pdf) return;
  E.addManual(st.entities, t, 'TERMO');
  $('manualTerm').value = '';
  computeBoxes(); renderList(); drawBoxes(); updateButtons();
}
function updateButtons() {
  const has = !!st.pdf && st.licensed;
  $('btnPdf').disabled = !has; $('btnText').disabled = !has || !st.lines.length; $('btnTable').disabled = !has || !st.entities.length;
}
function mode() { return document.querySelector('input[name=mode]:checked').value; }
const XX = 'XXXXX';
function tokenFor(en) { return mode() === 'label' ? en.label : XX; }

/* Texto de cada linha visual depois da substituição, sem os dados e sem o que estiver sob caixas manuais. */
function replacedLines(p) {
  const out = [];
  const manual = st.manual[p.n] || [];
  const under = (m) => manual.some((b) => { const cx = (m.x0 + m.x1) / 2, cy = (m.yt + m.yb2) / 2; return cx >= b.x && cx <= b.x + b.w && cy >= b.y && cy <= b.y + b.h; });
  for (const ln of p.lines) {
    const starts = new Map();
    for (const en of st.entities) { if (!en.selected) continue; for (const r of en.ranges || []) if (r.ln === ln) starts.set(r.a, { b: r.b, tok: tokenFor(en) }); }
    let seg = null, k = 0, lastManual = false;
    const push = () => { if (seg && seg.text.trim()) out.push(seg); seg = null; };
    while (k < ln.text.length) {
      const m = ln.map[k];
      if (m && seg && Math.abs(m.bl - seg.bl) > seg.fh * 0.5) push();
      if (m && !seg) seg = { x: m.x0, bl: m.bl, fh: m.fh, text: '' };
      const r = starts.get(k);
      if (r) { if (seg) seg.text += r.tok; k = r.b; lastManual = false; continue; }
      if (m && under(m)) { if (seg && !lastManual) seg.text += XX; lastManual = true; k++; continue; }
      lastManual = false;
      if (seg) seg.text += ln.text[k];
      k++;
    }
    push();
  }
  return out;
}

/* --------------------------------------------------------------- saídas */
async function exportPdf() {
  const btn = $('btnPdf'); btn.disabled = true;
  try {
    const out = await PDFLib.PDFDocument.create();
    out.setProducer('Cofre PDF | Ana Marta Soares'); out.setCreator('Cofre PDF');
    const font = await out.embedFont(PDFLib.StandardFonts.Helvetica);
    const m = mode();
    for (const p of st.pages) {
      status('outStatus', `A gerar a página ${p.n} de ${st.pages.length}…`, 'busy');
      const vp = p.page.getViewport({ scale: EXPORT_SCALE });
      const cv = document.createElement('canvas'); cv.width = Math.ceil(vp.width); cv.height = Math.ceil(vp.height);
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
      await p.page.render({ canvasContext: ctx, viewport: vp }).promise;
      const S = EXPORT_SCALE;
      for (const en of st.entities) {
        if (!en.selected) continue;
        for (const b of en.boxes || []) {
          if (b.page !== p.n) continue;
          const x = (b.x0 - 1) * S, y = (b.yt - 1) * S, w = (b.x1 - b.x0 + 2) * S, h = (b.yb - b.yt + 2) * S;
          if (m === 'redact') { ctx.fillStyle = '#000'; ctx.fillRect(x, y, w, h); }
          else if (m === 'xxxxx') {
            ctx.fillStyle = '#fff'; ctx.fillRect(x, y, w, h);
            let fs = h * 0.72; ctx.font = `${fs}px Arial`;
            const tw = ctx.measureText(XX).width;
            if (tw > w - 2) { fs = fs * (w - 2) / tw; ctx.font = `${fs}px Arial`; }
            ctx.fillStyle = '#1e1a18'; ctx.textBaseline = 'middle'; ctx.textAlign = 'left';
            ctx.fillText(XX, x + h * 0.18, y + h / 2);
          } else {
            ctx.fillStyle = '#fff'; ctx.fillRect(x, y, w, h);
            ctx.strokeStyle = '#1e1a18'; ctx.lineWidth = Math.max(1, S * 0.4); ctx.strokeRect(x, y, w, h);
            let fs = h * 0.7; ctx.font = `600 ${fs}px Arial`;
            const tw = ctx.measureText(en.label).width;
            if (tw > w - 4) { fs = fs * (w - 4) / tw; ctx.font = `600 ${fs}px Arial`; }
            ctx.fillStyle = '#1e1a18'; ctx.textBaseline = 'middle'; ctx.textAlign = 'center';
            ctx.fillText(en.label, x + w / 2, y + h / 2);
          }
        }
      }
      ctx.fillStyle = '#000';
      (st.manual[p.n] || []).forEach((b) => ctx.fillRect(b.x * S, b.y * S, b.w * S, b.h * S));
      const blob = await new Promise((r) => cv.toBlob(r, 'image/jpeg', 0.9));
      const img = await out.embedJpg(await blob.arrayBuffer());
      const pg = out.addPage([p.w, p.h]);
      pg.drawImage(img, { x: 0, y: 0, width: p.w, height: p.h });
      // camada de texto pesquisável, que contém apenas o texto já substituído
      for (const seg of replacedLines(p)) {
        const txt = seg.text.replace(/[^\x20-\x7E\u00A0-\u00FF€‘’“”–—…•]/g, ' ');
        try { pg.drawText(txt, { x: seg.x, y: p.h - seg.bl, size: Math.max(4, seg.fh), font, opacity: 0 }); }
        catch (e) { /* carácter não suportado: a linha fica só na imagem */ }
      }
    }
    const bytes = await out.save();
    download(new Blob([bytes], { type: 'application/pdf' }), baseName() + ' (anonimizado).pdf');
    status('outStatus', 'PDF anonimizado descarregado. Os dados substituídos já não existem no ficheiro.', 'ok');
  } catch (e) { status('outStatus', 'Não foi possível gerar o PDF: ' + (e.message || e), 'err'); }
  finally { btn.disabled = false; }
}
function exportText() {
  const plan = E.buildPlan(st.entities, { style: mode() === 'label' ? 'label' : 'x' });
  const paras = [{ text: baseName() + ' (texto pseudonimizado)', bold: true, size: 14, after: 200 }];
  for (const p of st.pages) {
    paras.push({ text: 'Página ' + p.n, bold: true, size: 11, color: '5C544F', after: 80 });
    for (const ln of p.lines) paras.push({ text: E.applyToText(ln.text, plan), size: 11, after: 40 });
  }
  paras.push({ text: 'Cofre PDF | Ana Marta Soares, Especialista em IA aplicada ao Direito', size: 9, color: '5C544F' });
  download(new Blob([D.buildDocx(paras)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), baseName() + ' (texto pseudonimizado).docx');
  status('outStatus', 'Texto pseudonimizado descarregado.', 'ok');
}
function exportTable() {
  const text = E.buildTable(st.entities, { document: st.name }).replace('Cofre de Anonimização para Word', 'Cofre PDF');
  const paras = text.split('\n').map((line, i) => i === 0 ? { text: line, bold: true, size: 14, after: 160 }
    : /^Esta tabela/.test(line) ? { text: line, italic: true, size: 10, after: 200 } : { text: line, size: 11, after: 60 });
  download(new Blob([D.buildDocx(paras)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), baseName() + ' (tabela de correspondencias).docx');
  status('outStatus', 'Tabela descarregada. Guarde-a separadamente e não a partilhe.', 'ok');
}

/* -------------------------------------------------------------- arranque */
let deferredInstall = null;
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferredInstall = e; $('btnInstall').hidden = false; });
$('btnInstall').addEventListener('click', async () => { if (!deferredInstall) return; deferredInstall.prompt(); await deferredInstall.userChoice; deferredInstall = null; $('btnInstall').hidden = true; });

$('fileInput').addEventListener('change', (e) => { openFile(e.target.files[0]); e.target.value = ''; });
const drop = $('drop');
['dragenter', 'dragover'].forEach((t) => document.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => document.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
document.addEventListener('drop', (e) => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) openFile(f); });
$('btnActivate').addEventListener('click', activate);
$('btnDeactivate').addEventListener('click', () => { storeSet(LIC_KEY, null); setLicensed(null); });
$('btnAll').addEventListener('click', () => { st.entities.forEach((e) => (e.selected = true)); renderList(); drawBoxes(); });
$('btnNone').addEventListener('click', () => { st.entities.forEach((e) => (e.selected = false)); renderList(); drawBoxes(); });
$('btnAddTerm').addEventListener('click', addTerm);
$('manualTerm').addEventListener('keydown', (e) => { if (e.key === 'Enter') addTerm(); });
$('btnPdf').addEventListener('click', exportPdf);
$('btnText').addEventListener('click', exportText);
$('btnTable').addEventListener('click', exportTable);

// abertura directa a partir do sistema ("Abrir com > Cofre PDF")
if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (params) => {
    if (params.files && params.files.length) openFile(await params.files[0].getFile());
  });
}
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
initLicense();
window.__cofrePdf = st;
