/*
 * Cofre Transcrição: interface.
 * Ana Marta Soares, Especialista em IA aplicada ao Direito.
 * A gravação é lida e transcrita neste computador; nada é enviado para servidores.
 */
import { openMedia, readSegment } from './decode.js';

const E = window.AnonEngine, D = window.DocxUtil, L = window.Licenca;
const $ = (id) => document.getElementById(id);
const LIC_KEY = 'cofre.audio.licenca';
const PIECE = 300; // segundos por parte
const COLORS = ['#1e1a18', '#8a5a2b', '#3f6b5a', '#6b4a7a', '#2f5d8a', '#9b3d3d', '#7a7a2b', '#5c544f'];

const st = { licensed: false, file: null, duration: 0, segs: [], names: [], running: false, stop: false, worker: null, gpu: null, loadedKey: '' };

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
  if (!L.allows('audio', v.tipo)) return { ok: false, motivo: 'Este código é de outro produto e não activa o Cofre Transcrição.' };
  if ((await revoked()).includes(v.num)) return { ok: false, motivo: 'Esta licença foi desactivada. Contacte Ana Marta Soares.' };
  return v;
}
function setLicensed(v) {
  st.licensed = !!(v && v.ok);
  $('licGate').hidden = st.licensed;
  $('panel').hidden = !st.licensed;
  $('licInfo').textContent = st.licensed ? `${v.tipoNome} n.º ${v.num}, ${v.validade ? 'válida até ' + v.validade.toLocaleDateString('pt-PT') : 'sem prazo de validade'}.` : '';
  if (st.licensed && st.file && !st.running && !st.segs.length) start();
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
function hms(t) {
  t = Math.max(0, Math.floor(t));
  const h = Math.floor(t / 3600), m = Math.floor((t % 3600) / 60), s = t % 60;
  return (h ? String(h).padStart(2, '0') + ':' : '') + String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0');
}
function dur(t) {
  if (t < 600) { const m = Math.floor(t / 60), s = Math.round(t % 60); return (m ? m + (m === 1 ? ' minuto' : ' minutos') + (s ? ' e ' : '') : '') + (s ? s + ' segundos' : ''); }
  t = Math.round(t / 60);
  if (t < 1) return 'menos de 1 minuto';
  const h = Math.floor(t / 60), m = t % 60;
  return (h ? h + (h === 1 ? ' hora' : ' horas') + (m ? ' e ' : '') : '') + (m ? m + (m === 1 ? ' minuto' : ' minutos') : '');
}
function download(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
function baseName() { return (st.file ? st.file.name : 'gravacao').replace(/\.[^.]+$/, ''); }
function spkName(i) { return (st.names[i] && st.names[i].trim()) || ('Orador ' + (i + 1)); }

/* ------------------------------------------------------ equipamento */
async function detectDevice() {
  if (navigator.gpu) {
    try { const a = await navigator.gpu.requestAdapter(); if (a) return { device: 'webgpu', fp16: a.features.has('shader-f16') }; } catch (e) { }
  }
  return { device: 'wasm', fp16: false };
}
function modelConfig() {
  const q = document.querySelector('input[name=quality]:checked').value;
  const g = st.gpu;
  if (g.device === 'webgpu') {
    if (q === 'max') return { model: 'onnx-community/whisper-large-v3-turbo', dtype: { encoder_model: 'q4', decoder_model_merged: 'q4' }, size: '760 MB' };
    return { model: 'onnx-community/whisper-small', dtype: { encoder_model: g.fp16 ? 'fp16' : 'fp32', decoder_model_merged: 'q4' }, size: g.fp16 ? '410 MB' : '590 MB' };
  }
  if (q === 'max') return { model: 'onnx-community/whisper-small', dtype: { encoder_model: 'q8', decoder_model_merged: 'q8' }, size: '250 MB' };
  return { model: 'onnx-community/whisper-base', dtype: { encoder_model: 'q8', decoder_model_merged: 'q8' }, size: '80 MB' };
}

/* ------------------------------------------------------ transcrição */
function worker() {
  if (!st.worker) st.worker = new Worker(new URL('./asr-worker.js', import.meta.url), { type: 'module' });
  return st.worker;
}
function call(msg, onProgress) {
  return new Promise((res, rej) => {
    const w = worker();
    w.onmessage = (e) => {
      const d = e.data;
      if (d.type === 'load-progress') { if (onProgress) onProgress(d); return; }
      if (d.type === 'error') rej(new Error(d.message)); else res(d);
    };
    w.postMessage(msg, msg.audio ? [msg.audio.buffer] : []);
  });
}
function progress(title, frac, eta, note) {
  $('progress').hidden = false;
  $('progTitle').textContent = title;
  $('progBar').style.width = Math.round(Math.max(0, Math.min(1, frac)) * 100) + '%';
  $('progEta').textContent = eta || '';
  if (note != null) $('progNote').textContent = note;
}

async function openFile(file) {
  if (!file || st.running) return;
  st.file = file; st.segs = []; st.names = [];
  $('fileName').textContent = file.name;
  $('empty').hidden = true; $('work').hidden = false;
  $('transcript').innerHTML = ''; $('outCard').hidden = true; $('spkCard').hidden = true;
  if (st.licensed) start();
  else progress('Active a licença para iniciar a transcrição.', 0, '', '');
}

async function start() {
  if (st.running || !st.file) return;
  st.running = true; st.stop = false; st.segs = [];
  $('btnStop').hidden = false; $('openLabel').classList.add('disabled');
  const cfg = modelConfig();
  const speakers = $('optSpeakers').checked;
  try {
    progress('A ler a gravação…', 0, '', '');
    const { path, duration } = await openMedia(st.file);
    st.duration = duration;
    renderTranscript();
    const key = cfg.model + JSON.stringify(cfg.dtype) + speakers;
    if (st.loadedKey !== key) {
      if (st.worker) { st.worker.terminate(); st.worker = null; }
      progress('A preparar o modelo de transcrição…', 0, '', `Na primeira utilização é descarregado o modelo (cerca de ${cfg.size}). Fica guardado neste computador.`);
      await call({ type: 'load', config: { model: cfg.model, dtype: cfg.dtype, device: st.gpu.device, speakers } }, (d) => {
        const mb = (d.loaded / 1e6).toFixed(0), tot = d.total ? (d.total / 1e6).toFixed(0) : '?';
        progress(`A descarregar o ${d.label}…`, d.total ? d.loaded / d.total : 0, `${mb} de ${tot} MB`);
      });
      st.loadedKey = key;
    }
    let prev = null; const t0 = performance.now(); let doneSec = 0;
    for (let off = 0; off < duration && !st.stop; off += PIECE) {
      const len = Math.min(PIECE, duration - off);
      const elapsed = (performance.now() - t0) / 1000;
      const eta = doneSec > 0 ? 'Tempo restante estimado: ' + dur((duration - doneSec) * (elapsed / doneSec)) : 'A calcular o tempo restante…';
      progress(`A transcrever: ${hms(off)} de ${hms(duration)}`, off / duration, eta, 'Mantenha esta janela aberta. Pode continuar a trabalhar noutras aplicações.');
      const audio = await readSegment(path, off, len);
      const r = await call({ type: 'chunk', id: off, audio, offset: off, prevSpeaker: prev });
      prev = r.last;
      for (const s of r.segs) { if (s.start < off + len + 0.5) st.segs.push(s); }
      doneSec = Math.min(duration, off + len);
      renderTranscript();
    }
    if (st.stop) progress('Transcrição interrompida em ' + hms(doneSec) + '.', doneSec / duration, '', 'Pode exportar o que já foi transcrito.');
    else progress('Transcrição concluída.', 1, 'Duração: ' + dur(duration), 'Reveja o texto e atribua os nomes aos oradores antes de exportar.');
  } catch (e) {
    progress('Não foi possível concluir a transcrição.', 0, '', (e && e.message) || String(e));
  } finally {
    st.running = false; $('btnStop').hidden = true; $('openLabel').classList.remove('disabled');
    renderTranscript();
  }
}

/* ------------------------------------------------------ apresentação */
function blocks(segs) {
  // agrupa por minuto e junta frases seguidas do mesmo orador
  const mins = new Map();
  for (const s of segs) {
    const m = Math.floor(s.start / 60);
    if (!mins.has(m)) mins.set(m, []);
    const arr = mins.get(m), last = arr[arr.length - 1];
    if (last && last.speaker === s.speaker) { last.text += ' ' + s.text; last.end = s.end; }
    else arr.push({ start: s.start, end: s.end, speaker: s.speaker, text: s.text });
  }
  return [...mins.entries()].sort((a, b) => a[0] - b[0]);
}
function renderTranscript() {
  const host = $('transcript');
  if (!st.segs.length) { host.innerHTML = ''; return; }
  const useSpk = $('optSpeakers').checked;
  const nSpk = st.segs.reduce((m, s) => Math.max(m, (s.speaker || 0) + 1), 0);
  let html = `<h2 class="doc-h">${esc(baseName())}</h2><p class="meta">Duração: ${dur(st.duration)}${useSpk ? ' · Oradores identificados: ' + nSpk : ''}</p>`;
  for (const [m, arr] of blocks(st.segs)) {
    html += `<section class="minute"><h3>MINUTO ${hms(m * 60)}</h3>`;
    for (const b of arr) {
      html += `<p class="seg"><span class="t">${hms(b.start)}</span>${useSpk ? `<span class="s" style="color:${COLORS[b.speaker % COLORS.length]}">${esc(spkName(b.speaker))}:</span>` : ''}${esc(b.text)}</p>`;
    }
    html += '</section>';
  }
  host.innerHTML = html;
  renderSpeakers(nSpk, useSpk);
  $('outCard').hidden = false;
}
function renderSpeakers(n, useSpk) {
  $('spkCard').hidden = !useSpk || !n;
  const list = $('spkList');
  if (list.children.length === n) return;
  list.innerHTML = '';
  for (let i = 0; i < n; i++) {
    const cnt = st.segs.filter((s) => s.speaker === i).length;
    const row = document.createElement('div'); row.className = 'spk';
    row.innerHTML = `<span class="dot" style="background:${COLORS[i % COLORS.length]}"></span><input type="text" value="${esc(st.names[i] || '')}" placeholder="Orador ${i + 1}"><span class="n">${cnt}</span>`;
    row.querySelector('input').addEventListener('input', (e) => { st.names[i] = e.target.value; renderTranscript(); });
    list.appendChild(row);
  }
}

/* ------------------------------------------------------------ exportação */
function exportLines() {
  const useSpk = $('optSpeakers').checked;
  // oradores com o mesmo nome são tratados como a mesma pessoa
  const lines = [];
  for (const [m, arr] of blocks(st.segs)) {
    lines.push({ kind: 'min', text: 'Minuto ' + hms(m * 60) });
    let last = null;
    for (const b of arr) {
      const who = useSpk ? spkName(b.speaker) : '';
      if (last && last.who === who) { last.text += ' ' + b.text; continue; }
      last = { kind: 'seg', t: hms(b.start), who, text: b.text };
      lines.push(last);
    }
  }
  if ($('optAnon').checked) {
    const res = E.analyze(lines.filter((l) => l.kind === 'seg').map((l) => l.text), {});
    const plan = E.buildPlan(res.entities, { style: 'x' });
    lines.forEach((l) => { if (l.kind === 'seg') l.text = E.applyToText(l.text, plan); });
  }
  return lines;
}
function exportDocx() {
  const lines = exportLines();
  const paras = [
    { text: baseName(), bold: true, size: 16, after: 60 },
    { text: 'Transcrição automática · Duração: ' + dur(st.duration) + ' · ' + new Date().toLocaleDateString('pt-PT'), size: 10, color: '5C544F', after: 240 }
  ];
  for (const l of lines) {
    if (l.kind === 'min') paras.push({ text: l.text, bold: true, size: 10, color: '5C544F', after: 80 });
    else paras.push({ text: '[' + l.t + ']  ' + (l.who ? l.who + ': ' : '') + l.text, size: 11, after: 100 });
  }
  paras.push({ text: 'Transcrição gerada pelo Cofre Transcrição, a rever antes de utilização. Ana Marta Soares, Especialista em IA aplicada ao Direito', size: 9, color: '5C544F' });
  download(new Blob([D.buildDocx(paras)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), baseName() + ' (transcricao).docx');
  status('outStatus', 'Transcrição descarregada em Word.', 'ok');
}
function exportTxt() {
  const lines = exportLines();
  const out = [baseName(), 'Transcrição automática · Duração: ' + dur(st.duration), ''];
  for (const l of lines) out.push(l.kind === 'min' ? '\n' + l.text.toUpperCase() : '[' + l.t + '] ' + (l.who ? l.who + ': ' : '') + l.text);
  out.push('', 'Cofre Transcrição | Ana Marta Soares, Especialista em IA aplicada ao Direito');
  download(new Blob([out.join('\n')], { type: 'text/plain;charset=utf-8' }), baseName() + ' (transcricao).txt');
  status('outStatus', 'Transcrição descarregada em texto.', 'ok');
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
$('btnStop').addEventListener('click', () => { st.stop = true; $('btnStop').hidden = true; progress('A interromper após a parte em curso…', parseFloat($('progBar').style.width) / 100, '', ''); });
$('btnDocx').addEventListener('click', exportDocx);
$('btnTxt').addEventListener('click', exportTxt);
$('optSpeakers').addEventListener('change', renderTranscript);
window.addEventListener('beforeunload', (e) => { if (st.running) { e.preventDefault(); e.returnValue = ''; } });

if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (params) => { if (params.files && params.files.length) openFile(await params.files[0].getFile()); });
}

(async () => {
  st.gpu = await detectDevice();
  $('deviceInfo').textContent = st.gpu.device === 'webgpu'
    ? 'A transcrição utiliza a placa gráfica deste computador.'
    : 'Este browser não permite utilizar a placa gráfica; a transcrição usa o processador e é mais lenta. Para maior rapidez, utilize o Google Chrome ou o Microsoft Edge actualizados.';
  initLicense();
})();
window.__cofreAudio = st;
