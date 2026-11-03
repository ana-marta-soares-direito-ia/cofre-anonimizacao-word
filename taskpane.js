/*
 * Cofre de Anonimização para Word: painel lateral.
 * Ana Marta Soares, Especialista em IA aplicada ao Direito.
 * O conteúdo do documento é lido e alterado através da API do Word e
 * processado localmente pelo motor (engine.js). Não há pedidos de rede.
 */
'use strict';
(function () {
  var E = window.AnonEngine, D = window.DocxUtil;
  var $ = function (id) { return document.getElementById(id); };

  var state = {
    entities: [],        // resultado da última análise
    applied: [],         // entidades já substituídas nesta sessão (para a tabela)
    locate: {},          // índice de localização por entidade
    marked: [],          // variantes realçadas para revisão
    docName: '',
    revertMap: [],
    licensed: false,
    tab: 'anon'
  };
  var LIC_KEY = 'cofre.licenca';

  /* ---------------------------------------------------------- licença */
  function storeGet() { try { return window.localStorage.getItem(LIC_KEY); } catch (e) { return null; } }
  function storeSet(v) { try { if (v) window.localStorage.setItem(LIC_KEY, v); else window.localStorage.removeItem(LIC_KEY); return true; } catch (e) { return false; } }
  function revokedList() {
    // lista de códigos bloqueados, publicada no mesmo endereço; sem ligação, a licença mantém-se
    return fetch('revogadas.json?t=' + Date.now(), { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : { revogadas: [] }; })
      .then(function (j) { return (j && j.revogadas) || []; })
      .catch(function () { return []; });
  }
  function checkCode(code) {
    var v = window.Licenca ? window.Licenca.verify(code) : { ok: false, motivo: 'Não foi possível verificar a licença.' };
    if (!v.ok) return Promise.resolve(v);
    if (window.Licenca.allows && !window.Licenca.allows('word', v.tipo)) return Promise.resolve({ ok: false, motivo: 'Este código é de outro produto (Cofre PDF) e não activa o suplemento do Word.' });
    return revokedList().then(function (list) {
      if (list.indexOf(v.num) >= 0) { v.ok = false; v.motivo = 'Esta licença foi desactivada. Contacte Ana Marta Soares.'; }
      return v;
    });
  }
  function setLicensed(v) {
    state.licensed = !!(v && v.ok);
    if (state.licensed) {
      $('licInfo').textContent = v.tipoNome + ' n.º ' + v.num + ', ' + (v.validade ? 'válida até ' + v.validade.toLocaleDateString('pt-PT') : 'sem prazo de validade') + '.';
      $('btnDeactivate').hidden = false;
    } else {
      $('licInfo').textContent = 'Sem licença activada.';
      $('btnDeactivate').hidden = true;
    }
    showTab(state.tab);
  }
  function initLicense() {
    var code = storeGet();
    if (!code) { setLicensed(null); return; }
    checkCode(code).then(function (v) {
      if (!v.ok) { storeSet(null); setStatus('licStatus', v.motivo, 'err'); }
      setLicensed(v);
    });
  }
  function activate() {
    var code = $('licCode').value.replace(/\s+/g, '');
    if (!code) { setStatus('licStatus', 'Cole o código de licença.', 'err'); return; }
    setStatus('licStatus', 'A verificar…', 'busy');
    checkCode(code).then(function (v) {
      if (!v.ok) { setStatus('licStatus', v.motivo, 'err'); return; }
      if (!storeSet(code)) setStatus('licStatus', 'Licença válida. Esta versão do Word não permite guardar a activação, pelo que poderá ser necessário repeti-la.', 'busy');
      else setStatus('licStatus', '');
      $('licCode').value = '';
      setLicensed(v);
    });
  }
  function deactivate() {
    storeSet(null);
    setLicensed(null);
    setStatus('licStatus', 'A licença foi removida deste computador.', 'ok');
  }
  function showTab(name) {
    state.tab = name;
    document.querySelectorAll('.tab').forEach(function (x) { var on = x.dataset.tab === name; x.classList.toggle('active', on); x.setAttribute('aria-selected', on); });
    var gate = !state.licensed && name !== 'info';
    $('licGate').hidden = !gate;
    document.querySelectorAll('.panel').forEach(function (p) { p.classList.toggle('active', !gate && p.id === 'tab-' + name); });
  }

  /* ---------------------------------------------------------- arranque */
  function boot() {
    if (!window.__cofreSuportado || !E || !D) { $('unsupported').hidden = false; disableAll(); return; }
    buildCategoryList();
    buildManualTypes();
    bindUI();
    initLicense();
    if (typeof Office === 'undefined') { $('outside').hidden = false; disableWord(); return; }
    Office.onReady(function (info) {
      if (!info || info.host !== Office.HostType.Word) { $('outside').hidden = false; disableWord(); return; }
      if (!supports('1.3')) { $('unsupported').hidden = false; disableAll(); return; }
      try {
        var url = Office.context.document && Office.context.document.url;
        if (url) state.docName = decodeURIComponent(String(url).split(/[\\/]/).pop());
      } catch (e) { /* sem nome */ }
    });
  }
  function supports(v) {
    try { return Office.context.requirements.isSetSupported('WordApi', v); } catch (e) { return false; }
  }
  function disableAll() { document.querySelectorAll('button, input, select, textarea').forEach(function (b) { if (!b.classList.contains('tab')) b.disabled = true; }); }
  function disableWord() { ['btnAnalyze', 'btnApply', 'btnRevert'].forEach(function (id) { $(id).disabled = true; }); }

  /* -------------------------------------------------------- interface */
  function bindUI() {
    document.querySelectorAll('.tab').forEach(function (t) {
      t.addEventListener('click', function () {
        showTab(t.dataset.tab);
        window.scrollTo(0, 0);
      });
    });
    $('btnAnalyze').addEventListener('click', function () { analyze(false); });
    $('btnRecheck').addEventListener('click', function () { analyze(true); });
    $('btnAll').addEventListener('click', function () { setAll(true); });
    $('btnNone').addEventListener('click', function () { setAll(false); });
    $('btnMark').addEventListener('click', markForReview);
    $('btnUnmark').addEventListener('click', unmark);
    $('btnAddTerm').addEventListener('click', addTerm);
    $('manualTerm').addEventListener('keydown', function (e) { if (e.key === 'Enter') addTerm(); });
    $('btnApply').addEventListener('click', askConfirm);
    $('btnCancel').addEventListener('click', function () { $('confirmApply').hidden = true; $('btnApply').hidden = false; });
    $('btnConfirm').addEventListener('click', apply);
    $('btnTableDoc').addEventListener('click', openTableDoc);
    $('btnTableCopy').addEventListener('click', copyTable);
    $('tableFile').addEventListener('change', readTableFile);
    $('revertText').addEventListener('input', previewRevert);
    $('btnRevert').addEventListener('click', revert);
    $('btnActivate').addEventListener('click', activate);
    $('btnDeactivate').addEventListener('click', deactivate);
  }

  function buildCategoryList() {
    var box = $('catList');
    E.TYPE_ORDER.forEach(function (t) {
      if (t === 'TERMO') return;
      var l = document.createElement('label');
      l.className = 'check';
      l.innerHTML = '<input type="checkbox" data-type="' + t + '" checked><span>' + esc(E.TYPES[t].name) + '</span>';
      box.appendChild(l);
    });
  }
  function buildManualTypes() {
    var sel = $('manualType');
    ['TERMO', 'PESSOA', 'ENTIDADE', 'MORADA'].forEach(function (t) {
      var o = document.createElement('option');
      o.value = t; o.textContent = { TERMO: 'Outro termo', PESSOA: 'Pessoa', ENTIDADE: 'Entidade', MORADA: 'Morada' }[t];
      sel.appendChild(o);
    });
  }
  function enabledTypes() {
    var r = {};
    document.querySelectorAll('#catList input').forEach(function (c) { r[c.dataset.type] = c.checked; });
    return r;
  }
  function setStatus(id, msg, kind) {
    var el = $(id);
    el.className = 'status' + (kind ? ' ' + kind : '');
    el.textContent = msg || '';
  }
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function plural(n, a, b) { return n + ' ' + (n === 1 ? a : b); }
  function busy(btn, on, label) {
    if (on) { btn.dataset.label = btn.textContent; btn.textContent = label; btn.disabled = true; }
    else { btn.textContent = btn.dataset.label || btn.textContent; btn.disabled = false; }
  }
  function errMsg(e) {
    var m = (e && (e.message || e.debugInfo && e.debugInfo.message)) || String(e);
    return 'Não foi possível concluir a operação: ' + m;
  }

  /* ---------------------------------------------- recolha do documento */
  // Corpo, notas de rodapé e notas finais podem ser tratados em conjunto;
  // cabeçalhos e rodapés são tratados um a um (podem estar ligados entre secções).
  async function collectContainers(ctx) {
    var doc = ctx.document;
    var main = [{ kind: 'Corpo do texto', body: doc.body }];
    var hf = [];
    var sections = doc.sections;
    sections.load('items');
    await ctx.sync();
    var types = ['Primary', 'FirstPage', 'EvenPages'];
    var cand = [];
    sections.items.forEach(function (sec) {
      types.forEach(function (t) {
        cand.push({ kind: 'Cabeçalho', getter: function () { return sec.getHeader(t); } });
        cand.push({ kind: 'Rodapé', getter: function () { return sec.getFooter(t); } });
      });
    });
    // tenta carregar todos de uma vez; se falhar, carrega um a um
    try {
      cand.forEach(function (c) { c.body = c.getter(); c.body.load('text'); });
      await ctx.sync();
      cand.forEach(function (c) { if (c.body.text && c.body.text.trim()) hf.push(c); });
    } catch (e) {
      for (var i = 0; i < cand.length; i++) {
        try {
          var b = cand[i].getter(); b.load('text'); await ctx.sync();
          if (b.text && b.text.trim()) { cand[i].body = b; hf.push(cand[i]); }
        } catch (e2) { /* inexistente */ }
      }
    }
    if (supports('1.5')) {
      try {
        var fn = doc.body.footnotes, en = doc.body.endnotes;
        fn.load('items'); en.load('items');
        await ctx.sync();
        fn.items.forEach(function (n) { main.push({ kind: 'Nota de rodapé', body: n.body }); });
        en.items.forEach(function (n) { main.push({ kind: 'Nota final', body: n.body }); });
      } catch (e) { /* sem notas ou API indisponível */ }
    }
    return { main: main, hf: hf, all: main.concat(hf) };
  }

  async function readParagraphs(ctx, containers) {
    var cols = containers.all.map(function (c) { var p = c.body.paragraphs; p.load('items/text'); return p; });
    await ctx.sync();
    var texts = [];
    cols.forEach(function (col) { col.items.forEach(function (p) { texts.push(p.text); }); });
    return texts;
  }

  /* ------------------------------------------------------------ análise */
  async function analyze(recheck) {
    if (!state.licensed) { showTab('anon'); return; }
    var btn = recheck ? $('btnRecheck') : $('btnAnalyze');
    busy(btn, true, 'A analisar…');
    setStatus('analyzeStatus', 'A ler o documento…', 'busy');
    try {
      var texts = await Word.run(async function (ctx) {
        var c = await collectContainers(ctx);
        return readParagraphs(ctx, c);
      });
      setStatus('analyzeStatus', 'A identificar os dados…', 'busy');
      await tick();
      var existing = E.existingLabels(texts);
      // a numeração continua a partir das etiquetas já existentes e já atribuídas
      state.applied.forEach(function (en) {
        var m = en.label.match(/^\[([A-Z_]+?)_(\d+)\]$/);
        if (m) existing[m[1]] = Math.max(existing[m[1]] || 0, +m[2]);
      });
      var res = E.analyze(texts, { types: enabledTypes(), existingLabels: existing });
      state.entities = res.entities;
      state.locate = {};
      renderEntities();
      $('review').hidden = false;
      $('applyCard').hidden = false;
      $('confirmApply').hidden = true; $('btnApply').hidden = false;
      setStatus('applyStatus', '');
      var n = state.entities.length;
      if (recheck) {
        setStatus('analyzeStatus', n ? 'Nova verificação concluída: ' + plural(n, 'elemento por tratar.', 'elementos por tratar.') : 'Nova verificação concluída: não foram identificados dados por tratar.', n ? 'busy' : 'ok');
        $('review').scrollIntoView({ behavior: 'smooth', block: 'start' });
      } else {
        setStatus('analyzeStatus', 'Análise concluída.', 'ok');
      }
    } catch (e) {
      setStatus('analyzeStatus', errMsg(e), 'err');
    } finally { busy(btn, false); }
  }
  function tick() { return new Promise(function (r) { setTimeout(r, 15); }); }

  function renderEntities() {
    var list = $('entityList');
    list.innerHTML = '';
    var ents = state.entities;
    var total = ents.reduce(function (a, e) { return a + e.count; }, 0);
    var groups = {};
    ents.forEach(function (en) { (groups[en.type] = groups[en.type] || []).push(en); });
    var nTypes = Object.keys(groups).length;
    $('reviewSummary').textContent = ents.length
      ? plural(ents.length, 'elemento distinto', 'elementos distintos') + ', em ' + plural(total, 'ocorrência', 'ocorrências') + ' e ' + plural(nTypes, 'categoria', 'categorias') + '. Desmarque o que não deve ser substituído e acrescente o que não tiver sido identificado.'
      : 'Não foram identificados dados pessoais. Pode acrescentar manualmente os termos a substituir.';
    if (!ents.length) { list.innerHTML = '<div class="empty">Sem elementos identificados.</div>'; return; }
    E.TYPE_ORDER.forEach(function (t) {
      var g = groups[t];
      if (!g) return;
      var wrap = document.createElement('div');
      wrap.className = 'group';
      var head = document.createElement('div');
      head.className = 'group-h';
      var allOn = g.every(function (e) { return e.selected; });
      head.innerHTML = '<span class="chev">▼</span><input type="checkbox" ' + (allOn ? 'checked' : '') + ' aria-label="Seleccionar a categoria"><span>' + esc(E.TYPES[t].name) + '</span><span class="cnt">' + g.length + '</span>';
      var gchk = head.querySelector('input');
      gchk.addEventListener('click', function (ev) { ev.stopPropagation(); g.forEach(function (e) { e.selected = gchk.checked; }); renderEntities(); });
      head.addEventListener('click', function () { wrap.classList.toggle('closed'); });
      wrap.appendChild(head);
      var items = document.createElement('div');
      items.className = 'items';
      g.forEach(function (en) {
        var row = document.createElement('div');
        row.className = 'item' + (en.selected ? '' : ' off');
        var vars = Array.from(en.variants.keys()).filter(function (v) { return v !== en.canonical; });
        row.innerHTML = '<input type="checkbox" ' + (en.selected ? 'checked' : '') + ' aria-label="Substituir">' +
          '<div class="txt"><div class="val">' + esc(en.canonical) + '</div>' +
          '<div class="meta"><span class="chip">' + esc(en.label) + '</span>' +
          (en.manual ? '<span class="times">acrescentado</span>' : '<span class="times">' + plural(en.count, 'ocorrência', 'ocorrências') + '</span>') +
          '</div>' + (vars.length ? '<div class="var">Também: ' + esc(vars.join('; ')) + '</div>' : '') + '</div>' +
          '<button class="loc" title="Localizar no documento">Localizar</button>';
        row.querySelector('input').addEventListener('change', function (ev) {
          en.selected = ev.target.checked; row.classList.toggle('off', !en.selected);
          gchk.checked = g.every(function (e) { return e.selected; });
        });
        row.querySelector('.loc').addEventListener('click', function () { locate(en); });
        items.appendChild(row);
      });
      wrap.appendChild(items);
      list.appendChild(wrap);
    });
  }
  function setAll(v) { state.entities.forEach(function (e) { e.selected = v; }); renderEntities(); }

  function addTerm() {
    var term = $('manualTerm').value.trim();
    if (!term) return;
    if (term.length > 250) { setStatus('analyzeStatus', 'O termo é demasiado extenso (máximo de 250 caracteres).', 'err'); return; }
    var en = E.addManual(state.entities.concat(state.applied), term, $('manualType').value);
    if (en && state.entities.indexOf(en) < 0) {
      if (state.applied.indexOf(en) >= 0) { en = null; }
      else state.entities.push(en);
    }
    $('manualTerm').value = '';
    renderEntities();
    $('review').hidden = false; $('applyCard').hidden = false;
  }

  /* --------------------------------------------- pesquisa no Word */
  function wordSafe(s) { return s.replace(/\^/g, '^^').replace(/[\n\u000b]/g, '^l').replace(/\t/g, '^t'); } // quebra de linha manual e tabulação na pesquisa do Word
  function searchOpts(it) { return { matchCase: !!it.matchCase, matchWholeWord: !!it.wholeWord }; }

  async function locate(en) {
    try {
      await Word.run(async function (ctx) {
        var c = await collectContainers(ctx);
        var cols = [];
        E.unitsOf(en).forEach(function (u) {
          var v = u.prefix + u.text;
          var it = { matchCase: !en.manual, wholeWord: E.isWordy(v) };
          c.all.forEach(function (ct) { var r = ct.body.search(wordSafe(v), searchOpts(it)); r.load('items'); cols.push(r); });
        });
        await ctx.sync();
        var ranges = [];
        cols.forEach(function (r) { r.items.forEach(function (x) { ranges.push(x); }); });
        if (!ranges.length) { setStatus('analyzeStatus', 'Este elemento já não se encontra no documento.', 'busy'); return; }
        var i = (state.locate[en.id] || 0) % ranges.length;
        state.locate[en.id] = i + 1;
        ranges[i].select();
        await ctx.sync();
        setStatus('analyzeStatus', 'Ocorrência ' + (i + 1) + ' de ' + ranges.length + ' de ' + en.label + '.', 'busy');
      });
    } catch (e) { setStatus('analyzeStatus', errMsg(e), 'err'); }
  }

  async function markForReview() {
    var plan = E.buildPlan(state.entities, {});
    var flat = [].concat.apply([], plan);
    if (!flat.length) return;
    busy($('btnMark'), true, 'A realçar…');
    try {
      var n = await Word.run(async function (ctx) {
        var c = await collectContainers(ctx);
        var cols = [];
        flat.forEach(function (it) { c.all.forEach(function (ct) { var r = ct.body.search(wordSafe(it.search), searchOpts(it)); r.load('items'); cols.push(r); }); });
        await ctx.sync();
        var k = 0;
        cols.forEach(function (r) { r.items.forEach(function (x) { x.font.highlightColor = '#FFF2A8'; k++; }); });
        await ctx.sync();
        return k;
      });
      state.marked = flat;
      setStatus('analyzeStatus', plural(n, 'ocorrência realçada', 'ocorrências realçadas') + ' no documento.', 'ok');
    } catch (e) { setStatus('analyzeStatus', errMsg(e), 'err'); }
    finally { busy($('btnMark'), false); }
  }
  async function unmark() {
    var flat = state.marked.slice();
    // após a substituição, o realce encontra-se nas etiquetas
    state.applied.forEach(function (en) { flat.push({ search: en.label, matchCase: true, wholeWord: false }); });
    if (!flat.length) { setStatus('analyzeStatus', 'Não há realces aplicados pelo Cofre nesta sessão.', 'busy'); return; }
    busy($('btnUnmark'), true, 'A remover…');
    try {
      await Word.run(async function (ctx) {
        var c = await collectContainers(ctx);
        var cols = [];
        flat.forEach(function (it) { c.all.forEach(function (ct) { var r = ct.body.search(wordSafe(it.search), searchOpts(it)); r.load('items'); cols.push(r); }); });
        await ctx.sync();
        cols.forEach(function (r) { r.items.forEach(function (x) { x.font.highlightColor = null; }); });
        await ctx.sync();
      });
      state.marked = [];
      setStatus('analyzeStatus', 'Realce removido.', 'ok');
    } catch (e) { setStatus('analyzeStatus', errMsg(e), 'err'); }
    finally { busy($('btnUnmark'), false); }
  }

  /* ------------------------------------------------------- aplicação */
  function mode() { return document.querySelector('input[name=mode]:checked').value; }

  function askConfirm() {
    var sel = state.entities.filter(function (e) { return e.selected; });
    if (!sel.length) { setStatus('applyStatus', 'Não há elementos seleccionados.', 'err'); return; }
    var m = mode();
    var label = m === 'pseudo' ? 'pseudonimizados' : (m === 'anon' ? 'anonimizados' : 'rasurados');
    $('confirmText').textContent = 'Vão ser ' + label + ' ' + plural(sel.length, 'elemento', 'elementos') + ' no documento aberto. ' +
      (m === 'pseudo' ? 'A tabela de correspondências será apresentada no fim.' : 'Esta operação não permite a reposição dos dados através do Cofre.') + ' Pretende prosseguir?';
    $('btnApply').hidden = true;
    $('confirmApply').hidden = false;
  }

  async function apply() {
    if (!state.licensed) { showTab('anon'); return; }
    $('confirmApply').hidden = true; $('btnApply').hidden = false;
    var m = mode();
    var plan = E.buildPlan(state.entities, { style: m === 'redact' ? 'redact' : 'label' });
    var opts = { revisions: $('optRevisions').checked, props: $('optProps').checked, highlight: $('optHighlight').checked };
    var btn = $('btnApply');
    busy(btn, true, 'A aplicar…');
    setStatus('applyStatus', 'A substituir os dados no documento…', 'busy');
    var t0 = Date.now();
    try {
      var r = await Word.run(async function (ctx) {
        var doc = ctx.document, notes = [];
        var prevMode = null;
        // 1. desactiva o controlo de alterações (as substituições não podem ficar como revisões)
        if (supports('1.4')) {
          try {
            doc.load('changeTrackingMode'); await ctx.sync();
            prevMode = doc.changeTrackingMode;
            if (prevMode && prevMode !== 'Off') { doc.changeTrackingMode = 'Off'; await ctx.sync(); }
          } catch (e) { prevMode = null; }
        }
        // 2. revisões pendentes
        if (opts.revisions) {
          if (supports('1.6')) {
            try { doc.body.getTrackedChanges().acceptAll(); await ctx.sync(); }
            catch (e) { notes.push('Não foi possível aceitar automaticamente as revisões pendentes. Verifique-as em Rever > Aceitar.'); }
          } else notes.push('Esta versão do Word não permite aceitar as revisões a partir do painel. Verifique-as em Rever > Aceitar.');
        }
        var c = await collectContainers(ctx);
        var counts = {}, total = 0, skipped = [];
        async function runOn(containers) {
          for (var li = 0; li < plan.length; li++) {
            var jobs = [];
            plan[li].forEach(function (it) {
              if (it.search.length > 255) { if (skipped.indexOf(it.search) < 0) skipped.push(it.search); return; }
              containers.forEach(function (ct) {
                var r = ct.body.search(wordSafe(it.search), searchOpts(it));
                r.load('items'); jobs.push({ it: it, r: r });
              });
            });
            if (!jobs.length) continue;
            await ctx.sync();
            var any = false;
            jobs.forEach(function (j) {
              j.r.items.forEach(function (rg) {
                var nr = rg.insertText(j.it.replacement, 'Replace');
                if (opts.highlight) nr.font.highlightColor = '#FFF2A8';
                counts[j.it.entityId] = (counts[j.it.entityId] || 0) + 1; total++; any = true;
              });
            });
            if (any) await ctx.sync();
          }
        }
        await runOn(c.main);
        for (var h = 0; h < c.hf.length; h++) await runOn([c.hf[h]]);
        // 3. comentários
        if (opts.revisions && supports('1.4')) {
          try {
            var cs = doc.body.getComments(); cs.load('items'); await ctx.sync();
            cs.items.forEach(function (x) { x.delete(); });
            if (cs.items.length) await ctx.sync();
          } catch (e) { notes.push('Não foi possível eliminar os comentários. Verifique-os em Rever > Eliminar.'); }
        }
        // 4. propriedades
        if (opts.props) {
          try {
            var p = doc.properties;
            ['author', 'title', 'subject', 'keywords', 'comments', 'company', 'manager', 'category'].forEach(function (k) { try { p[k] = ''; } catch (e) { } });
            await ctx.sync();
          } catch (e) { notes.push('Não foi possível limpar todas as propriedades do documento. Verifique-as em Ficheiro > Informações.'); }
        }
        // 5. repõe o controlo de alterações
        if (prevMode && prevMode !== 'Off') { try { doc.changeTrackingMode = prevMode; await ctx.sync(); } catch (e) { } }
        // 6. verificação: procura as variantes originais que possam subsistir
        var left = 0, leftItems = [];
        var vjobs = [];
        [].concat.apply([], plan).forEach(function (it) {
          if (it.search.length > 255) return;
          c.all.forEach(function (ct) { var r = ct.body.search(wordSafe(it.search), searchOpts(it)); r.load('items'); vjobs.push({ it: it, r: r }); });
        });
        if (vjobs.length) {
          await ctx.sync();
          vjobs.forEach(function (j) { if (j.r.items.length) { left += j.r.items.length; if (leftItems.indexOf(j.it.search) < 0) leftItems.push(j.it.search); } });
        }
        return { counts: counts, total: total, skipped: skipped, notes: notes, left: left, leftItems: leftItems };
      });
      var done = state.entities.filter(function (e) { return e.selected; });
      if (m === 'pseudo') done.forEach(function (en) { if (state.applied.indexOf(en) < 0) state.applied.push(en); });
      state.marked = [];
      showResult(r, m, done.length, (Date.now() - t0) / 1000);
      setStatus('applyStatus', '');
      // a lista passa a conter apenas o que ficou por tratar
      state.entities = state.entities.filter(function (e) { return !e.selected; });
      renderEntities();
      if (!state.entities.length) { $('review').hidden = true; $('applyCard').hidden = true; }
    } catch (e) {
      setStatus('applyStatus', errMsg(e), 'err');
    } finally { busy(btn, false); }
  }

  function showResult(r, m, nEnt, secs) {
    var html = '<div class="stats">' +
      '<span>Elementos tratados</span><span class="v">' + nEnt + '</span>' +
      '<span>Substituições efectuadas</span><span class="v">' + r.total + '</span>' +
      '<span>Ocorrências remanescentes</span><span class="v ' + (r.left ? 'bad' : 'good') + '">' + r.left + '</span>' +
      '<span>Tempo</span><span class="v">' + secs.toFixed(1).replace('.', ',') + ' s</span></div>';
    if (r.left) html += '<div class="notice warn small">Subsistem ocorrências de: ' + esc(r.leftItems.slice(0, 8).join('; ')) + (r.leftItems.length > 8 ? '…' : '') + '. Podem encontrar-se em caixas de texto ou noutros objectos que o painel não altera. Utilize "Localizar" ou a pesquisa do Word.</div>';
    if (r.skipped.length) html += '<div class="notice warn small">Não foi possível tratar ' + plural(r.skipped.length, 'termo', 'termos') + ' com mais de 255 caracteres.</div>';
    r.notes.forEach(function (n) { html += '<div class="notice warn small">' + esc(n) + '</div>'; });
    html += '<p class="muted small">Guarde o documento com outro nome antes de o partilhar ou de o trabalhar num sistema de inteligência artificial. Recomenda-se ainda uma leitura final, para verificar identificadores indirectos que subsistam no texto.</p>';
    $('resultSummary').innerHTML = html;
    if (m === 'pseudo' && state.applied.length) {
      $('tableText').value = E.buildTable(state.applied, { document: state.docName });
      $('tableBlock').hidden = false;
    } else if (!state.applied.length) {
      $('tableBlock').hidden = true;
    }
    setStatus('tableStatus', '');
    $('resultCard').hidden = false;
    $('resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  /* ------------------------------------------------ tabela */
  function tableParagraphs(text) {
    return text.split('\n').map(function (line, i) {
      if (i === 0) return { text: line, bold: true, size: 14, color: '4A2580', after: 160 };
      if (/^Esta tabela/.test(line)) return { text: line, italic: true, size: 10, color: '8A5A00', after: 200 };
      if (/^Cofre de Anonimização/.test(line)) return { text: line, size: 9, color: '4A2580' };
      return { text: line, size: 11, after: 60 };
    });
  }
  async function openTableDoc() {
    var text = $('tableText').value;
    if (!text) return;
    setStatus('tableStatus', 'A criar o documento…', 'busy');
    try {
      var b64 = D.toBase64(D.buildDocx(tableParagraphs(text)));
      await Word.run(async function (ctx) {
        var nd = ctx.application.createDocument(b64);
        await ctx.sync();
        nd.open();
        await ctx.sync();
      });
      setStatus('tableStatus', 'A tabela foi aberta num novo documento. Guarde-o num local separado do documento pseudonimizado.', 'ok');
    } catch (e) {
      setStatus('tableStatus', 'Não foi possível abrir um novo documento nesta versão do Word. Utilize o botão "Copiar" e cole a tabela num documento novo.', 'err');
    }
  }
  async function copyTable() {
    var ta = $('tableText'), ok = false;
    try { if (navigator.clipboard && navigator.clipboard.writeText) { await navigator.clipboard.writeText(ta.value); ok = true; } } catch (e) { ok = false; }
    if (!ok) {
      try { ta.removeAttribute('readonly'); ta.select(); ok = document.execCommand('copy'); ta.setAttribute('readonly', ''); window.getSelection().removeAllRanges(); } catch (e) { ok = false; }
    }
    setStatus('tableStatus', ok ? 'Tabela copiada para a área de transferência.' : 'Não foi possível copiar automaticamente. Seleccione o texto da caixa e utilize Ctrl+C (ou Cmd+C).', ok ? 'ok' : 'err');
  }

  /* ------------------------------------------------ reposição */
  function readTableFile(ev) {
    var f = ev.target.files && ev.target.files[0];
    if (!f) return;
    var fr = new FileReader();
    if (/\.docx$/i.test(f.name)) {
      fr.onload = function () {
        D.readDocxText(fr.result).then(function (t) { $('revertText').value = t; previewRevert(); })
          .catch(function (e) { setStatus('revertStatus', e.message, 'err'); });
      };
      fr.readAsArrayBuffer(f);
    } else {
      fr.onload = function () { $('revertText').value = fr.result; previewRevert(); };
      fr.readAsText(f, 'utf-8');
    }
    ev.target.value = '';
  }
  function previewRevert() {
    state.revertMap = E.parseTable($('revertText').value);
    var n = state.revertMap.length;
    $('revertPreview').textContent = $('revertText').value.trim()
      ? (n ? plural(n, 'correspondência reconhecida.', 'correspondências reconhecidas.') : 'Não foi reconhecida nenhuma linha no formato [ETIQUETA_N] = dado original.')
      : '';
    setStatus('revertStatus', '');
  }
  async function revert() {
    if (!state.licensed) { showTab('repor'); return; }
    previewRevert();
    var map = state.revertMap;
    if (!map.length) { setStatus('revertStatus', 'Indique a tabela de correspondências.', 'err'); return; }
    var btn = $('btnRevert');
    busy(btn, true, 'A repor…');
    setStatus('revertStatus', 'A repor os dados no documento…', 'busy');
    try {
      var n = await Word.run(async function (ctx) {
        var doc = ctx.document, prevMode = null;
        if (supports('1.4')) {
          try { doc.load('changeTrackingMode'); await ctx.sync(); prevMode = doc.changeTrackingMode; if (prevMode !== 'Off') { doc.changeTrackingMode = 'Off'; await ctx.sync(); } } catch (e) { prevMode = null; }
        }
        var c = await collectContainers(ctx), total = 0;
        async function pass(containers, bare) {
          var jobs = [];
          map.forEach(function (r) {
            var s = bare ? r.bare : r.label;
            containers.forEach(function (ct) { var q = ct.body.search(wordSafe(s), { matchCase: true, matchWholeWord: bare }); q.load('items'); jobs.push({ r: r, q: q }); });
          });
          await ctx.sync();
          var any = false;
          jobs.forEach(function (j) { j.q.items.forEach(function (rg) { var nr = rg.insertText(j.r.original, 'Replace'); nr.font.highlightColor = null; total++; any = true; }); });
          if (any) await ctx.sync();
        }
        await pass(c.main, false); await pass(c.main, true);
        for (var h = 0; h < c.hf.length; h++) { await pass([c.hf[h]], false); await pass([c.hf[h]], true); }
        if (prevMode && prevMode !== 'Off') { try { doc.changeTrackingMode = prevMode; await ctx.sync(); } catch (e) { } }
        return total;
      });
      setStatus('revertStatus', n ? plural(n, 'etiqueta reposta', 'etiquetas repostas') + ' com os dados originais.' : 'Não foram encontradas no documento etiquetas correspondentes à tabela indicada.', n ? 'ok' : 'err');
    } catch (e) { setStatus('revertStatus', errMsg(e), 'err'); }
    finally { busy(btn, false); }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
