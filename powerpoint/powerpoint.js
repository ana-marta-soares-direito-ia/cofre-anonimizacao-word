/*
 * Cofre para PowerPoint: painel lateral.
 * Assistente com instruções (melhorar textos, criar diapositivos, uniformizar o estilo, inserir gráficos,
 * tabelas e caixas de texto, com desfazer) e anonimização ou pseudonimização do texto dos diapositivos.
 * Ana Marta Soares, Especialista em IA aplicada ao Direito.
 */
'use strict';
(function () {
  var C = window.Cofre, E = window.AnonEngine, $ = C.$;
  var st = { plano: null, textos: [], entities: [], applied: [], nome: '' };
  var LARG = 960, ALT = 540; // dimensões habituais de um diapositivo 16:9, em pontos

  /* ============================================================ arranque */
  function boot() {
    if (!window.__cofreSuportado || !E || !C) { $('unsupported').hidden = false; return; }
    C.iniciar({ produto: 'powerpoint', nome: 'Cofre para PowerPoint' });
    bind();
    if (typeof Office === 'undefined') { $('outside').hidden = false; return; }
    Office.onReady(function (info) {
      if (!info || info.host !== Office.HostType.PowerPoint) { $('outside').hidden = false; return; }
      if (!suporta('1.4')) { $('unsupported').hidden = false; return; }
      try { var u = Office.context.document && Office.context.document.url; if (u) st.nome = decodeURIComponent(String(u).split(/[\\/]/).pop()); } catch (e) { /* sem nome */ }
    });
  }
  function suporta(v) { try { return Office.context.requirements.isSetSupported('PowerPointApi', v); } catch (e) { return false; } }
  function bind() {
    $('btnExecutar').addEventListener('click', preparar);
    $('btnAplicar').addEventListener('click', aplicarPlano);
    $('btnDescartar').addEventListener('click', function () { st.plano = null; $('planoCard').hidden = true; });
    $('btnDesfazer').addEventListener('click', desfazer);
    $('instrucao').addEventListener('keydown', function (e) { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) preparar(); });
    $('btnAnalyze').addEventListener('click', analisar);
    $('btnAddTerm').addEventListener('click', acrescentarTermo);
    $('btnApply').addEventListener('click', anonimizar);
    $('btnTableCopy').addEventListener('click', copiarTabela);
    $('btnTableDocx').addEventListener('click', function () { C.tabelaDocx('tabela de correspondencias.docx', $('tableText').value); });
    $('tableFile').addEventListener('change', function () { C.lerFicheiroTabela($('tableFile').files[0]).then(function (t) { $('revertText').value = t; }); });
    $('btnRevert').addEventListener('click', repor);
  }

  /* ============================================================ leitura */
  // Lê os diapositivos e as formas com texto. Devolve [{n, id, formas:[{id, nome, tipo, x, y, w, h, texto, fonte, tamanho}]}]
  async function lerDiapositivos() {
    return PowerPoint.run(async function (ctx) {
      var slides = ctx.presentation.slides; slides.load('items/id');
      await ctx.sync();
      var formas = slides.items.map(function (s) { var sh = s.shapes; sh.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height'); return sh; });
      await ctx.sync();
      var tr = [];
      formas.forEach(function (sh, i) {
        sh.items.forEach(function (f) {
          if (['GeometricShape', 'TextBox', 'Placeholder', 'Unsupported'].indexOf(f.type) < 0) return;
          var t = f.textFrame.textRange; t.load('text'); t.font.load('name,size');
          tr.push({ i: i, f: f, t: t });
        });
      });
      // as formas sem texto podem falhar ao carregar: tenta tudo de uma vez e, se falhar, uma a uma
      try { await ctx.sync(); }
      catch (e) {
        var ok = [];
        for (var k = 0; k < tr.length; k++) {
          try { var t2 = tr[k].f.textFrame.textRange; t2.load('text'); t2.font.load('name,size'); await ctx.sync(); tr[k].t = t2; ok.push(tr[k]); } catch (e2) { /* sem texto */ }
        }
        tr = ok;
      }
      var res = slides.items.map(function (s, i) { return { n: i + 1, id: s.id, formas: [] }; });
      tr.forEach(function (x) {
        res[x.i].formas.push({ id: x.f.id, nome: x.f.name, tipo: x.f.type, x: Math.round(x.f.left), y: Math.round(x.f.top), w: Math.round(x.f.width), h: Math.round(x.f.height), texto: x.t.text || '', fonte: x.t.font.name, tamanho: x.t.font.size });
      });
      return res;
    });
  }
  async function seleccionados() {
    if (!suporta('1.5')) return [];
    return PowerPoint.run(async function (ctx) { var s = ctx.presentation.getSelectedSlides(); s.load('items/id'); await ctx.sync(); return s.items.map(function (x) { return x.id; }); });
  }
  function eTitulo(f, formas) {
    var nome = String(f.nome || '');
    if (/sub\s*t[ií]tulo|subtitle/i.test(nome)) return false;
    if (/(^|\s)(t[ií]tulo|title)(\s|\d|$)/i.test(nome)) return true;
    var comTexto = formas.filter(function (x) { return x.texto.trim(); });
    var topo = comTexto.slice().sort(function (a, b) { return a.y - b.y; })[0];
    return topo === f && f.h < ALT * 0.3 && (f.tamanho || 0) >= Math.max.apply(null, comTexto.map(function (x) { return x.tamanho || 0; }));
  }

  /* ============================================================ assistente */
  var SISTEMA = [
    'És o assistente do Cofre para PowerPoint, ao serviço de juristas, advogados, formadores e serviços públicos em Portugal. Recebes uma instrução e o conteúdo de uma apresentação: os diapositivos, numerados a partir de 1, e as formas com texto de cada um (identificador, nome, posição e tamanho em pontos, texto).',
    'O diapositivo tem ' + LARG + ' por ' + ALT + ' pontos.',
    'Respondes SEMPRE e APENAS com um objecto JSON, sem texto antes ou depois, com esta forma:',
    '{"resposta":"texto curto em português europeu para o utilizador (o que vais fazer, ou a resposta se a instrução for uma pergunta)","accoes":[...]}',
    'Acções possíveis (usa só as necessárias, no máximo 20, pela ordem de execução):',
    '{"tipo":"texto","diapositivo":2,"forma":"identificador da forma","texto":"novo texto; separa os parágrafos com \\n"}  (substitui o texto de uma forma existente; mantém o número aproximado de linhas para o texto caber)',
    '{"tipo":"novoDiapositivo","depoisDe":3,"titulo":"...","pontos":["...","..."]}  (cria um diapositivo de título e conteúdo; no máximo 6 pontos curtos)',
    '{"tipo":"estilo","diapositivos":"todos","fonte":"Calibri","corTitulos":"#1E1A18","corTexto":"#3A3530","tamanhoTitulos":32,"tamanhoTexto":18}  (uniformiza o tipo de letra, as cores e os tamanhos; "diapositivos" pode ser "todos" ou uma lista como [1,2])',
    '{"tipo":"grafico","diapositivo":2,"grafico":"colunas|barras|linhas|circular","titulo":"...","rotulos":["2024","2025"],"series":[{"nome":"Processos","valores":[120,150]}],"posicao":"direita|centro|baixo"}  (insere um gráfico como imagem; usa só números que constem da apresentação ou da instrução)',
    '{"tipo":"tabela","diapositivo":2,"linhas":[["Cabeçalho 1","Cabeçalho 2"],["...","..."]],"posicao":"centro|baixo|direita"}',
    '{"tipo":"caixaTexto","diapositivo":2,"texto":"...","posicao":"topo|baixo|direita|centro"}',
    'Regras:',
    '1. Altera apenas o que a instrução pede. Para melhorar a redacção, mantém o sentido, os factos, os números e as referências legais, e não acrescentes informação nova.',
    '2. As etiquetas como [PESSOA_1] ou [NIF_1] substituem dados pessoais: mantém-nas exactamente como estão.',
    '3. Escreve em português europeu (grafia anterior ao Acordo Ortográfico de 1990), em registo profissional e jurídico, sem travessões.',
    '4. Usa apenas identificadores de formas e números de diapositivos que existam nos dados fornecidos.',
    '5. Se a instrução for apenas uma pergunta, responde em "resposta" e devolve "accoes":[].'
  ].join('\n');

  function textoContexto(slides, m, alvo) {
    return slides.filter(function (s) { return !alvo || alvo.indexOf(s.id) >= 0; }).map(function (s) {
      var linhas = s.formas.filter(function (f) { return f.texto.trim(); }).map(function (f) {
        return '  forma ' + f.id + ' «' + m.tapar(f.nome || '') + '»' + (eTitulo(f, s.formas) ? ' (título)' : '') + ' x=' + f.x + ' y=' + f.y + ' l=' + f.w + ' a=' + f.h + (f.tamanho ? ' letra ' + f.fonte + ' ' + f.tamanho : '') + ':\n    ' + m.tapar(f.texto).replace(/\r\n?|\n|\v/g, '\n    ');
      });
      return 'DIAPOSITIVO ' + s.n + (linhas.length ? '\n' + linhas.join('\n') : ' (sem texto)');
    }).join('\n');
  }
  async function preparar() {
    if (!C.licenciado()) return;
    var instr = $('instrucao').value.trim();
    if (!instr) { C.setStatus('assistStatus', 'Escreva o que pretende.', 'err'); return; }
    var btn = $('btnExecutar'); C.busy(btn, true, 'A preparar…');
    $('planoCard').hidden = true;
    try {
      C.setStatus('assistStatus', 'A ler a apresentação…', 'busy');
      var slides = await lerDiapositivos();
      var alvo = $('ambito').value === 'seleccao' ? await seleccionados() : null;
      if (alvo && !alvo.length) alvo = null;
      var textos = [instr];
      slides.forEach(function (s) { s.formas.forEach(function (f) { textos.push(f.texto); }); });
      var m = C.mascara(textos, $('optMascara').checked);
      C.setStatus('assistStatus', (m.n ? 'Dados pessoais substituídos por etiquetas antes do envio: ' + m.n + '. ' : '') + 'A preparar o plano…', 'busy');
      var txt = await C.pedirIA({ sistema: SISTEMA, mensagem: 'INSTRUÇÃO: ' + m.tapar(instr) + '\n\nAPRESENTAÇÃO (' + slides.length + ' diapositivos' + (alvo ? '; o utilizador seleccionou os diapositivos ' + slides.filter(function (s) { return alvo.indexOf(s.id) >= 0; }).map(function (s) { return s.n; }).join(', ') : '') + '):\n' + textoContexto(slides, m, alvo), maxTokens: 6000, esforco: 'medio' });
      var j = C.destaparFundo(C.extrairJSON(txt), m);
      var accoes = (Array.isArray(j.accoes) ? j.accoes : []).filter(function (a) { return a && DESCR[a.tipo]; }).slice(0, 20);
      st.plano = { accoes: accoes, slides: slides };
      mostrarPlano(j.resposta || '', accoes, slides);
      C.setStatus('assistStatus', '');
    } catch (e) { C.setStatus('assistStatus', e.message || C.errMsg(e), 'err'); }
    finally { C.busy(btn, false); }
  }
  var NOMES_GRAF = { colunas: 'colunas', barras: 'barras', linhas: 'linhas', circular: 'circular' };
  var DESCR = {
    texto: function (a) { return 'Diapositivo ' + a.diapositivo + ': novo texto «' + resumir(a.texto) + '».'; },
    novoDiapositivo: function (a) { return 'Novo diapositivo depois do ' + (a.depoisDe || 'último') + ': «' + resumir(a.titulo) + '», com ' + C.plural((a.pontos || []).length, 'ponto', 'pontos') + '.'; },
    estilo: function (a) { return 'Estilo em ' + (Array.isArray(a.diapositivos) ? 'os diapositivos ' + a.diapositivos.join(', ') : 'todos os diapositivos') + [a.fonte ? 'letra ' + a.fonte : '', a.tamanhoTitulos ? 'títulos ' + a.tamanhoTitulos : '', a.tamanhoTexto ? 'texto ' + a.tamanhoTexto : '', a.corTitulos ? 'cor dos títulos ' + a.corTitulos : '', a.corTexto ? 'cor do texto ' + a.corTexto : ''].filter(Boolean).map(function (x, i) { return (i ? ', ' : ': ') + x; }).join('') + '.'; },
    grafico: function (a) { return 'Diapositivo ' + a.diapositivo + ': gráfico de ' + (NOMES_GRAF[a.grafico] || 'colunas') + (a.titulo ? ' «' + resumir(a.titulo) + '»' : '') + '.'; },
    tabela: function (a) { return 'Diapositivo ' + a.diapositivo + ': tabela com ' + C.plural((a.linhas || []).length, 'linha', 'linhas') + '.'; },
    caixaTexto: function (a) { return 'Diapositivo ' + a.diapositivo + ': caixa de texto «' + resumir(a.texto) + '».'; }
  };
  function resumir(t) { t = String(t || '').replace(/\s+/g, ' ').trim(); return t.length > 90 ? t.slice(0, 87) + '…' : t; }
  function mostrarPlano(resp, accoes) {
    $('resposta').textContent = resp;
    var ol = $('plano'); ol.innerHTML = '';
    accoes.forEach(function (a) { var li = document.createElement('li'); var d; try { d = DESCR[a.tipo](a); } catch (e) { d = a.tipo; } li.textContent = d; ol.appendChild(li); });
    $('btnAplicar').hidden = !accoes.length;
    $('btnDescartar').textContent = accoes.length ? 'Descartar' : 'Fechar';
    C.setStatus('aplicarStatus', '');
    $('planoCard').hidden = false;
    $('planoCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  async function aplicarPlano() {
    if (!st.plano) return;
    var btn = $('btnAplicar'); C.busy(btn, true, 'A aplicar…');
    var itens = $('plano').children, ok = 0, falhas = [];
    // os diapositivos novos são criados no fim, para não alterar a numeração usada nas outras acções
    var ordem = st.plano.accoes.map(function (a, i) { return i; }).sort(function (x, y) { return (st.plano.accoes[x].tipo === 'novoDiapositivo') - (st.plano.accoes[y].tipo === 'novoDiapositivo') || x - y; });
    var ids = st.plano.slides.map(function (s) { return s.id; });
    for (var k = 0; k < ordem.length; k++) {
      var i = ordem[k], a = st.plano.accoes[i];
      try { await executar(a, ids, st.plano.slides); ok++; itens[i].classList.add('feito'); }
      catch (e) { falhas.push((i + 1) + ': ' + (e.message || e)); itens[i].classList.add('falhou'); itens[i].title = e.message || String(e); }
    }
    C.busy(btn, false); btn.hidden = true;
    $('btnDescartar').textContent = 'Fechar';
    C.setStatus('aplicarStatus', falhas.length ? 'Aplicadas ' + ok + ' de ' + st.plano.accoes.length + ' acções. Não foi possível aplicar: ' + falhas.join('; ') : 'Plano aplicado. Pode desfazer as alterações com o botão abaixo.', falhas.length ? 'err' : 'ok');
    st.plano = null;
  }
  async function desfazer() {
    var btn = $('btnDesfazer'); C.busy(btn, true, 'A desfazer…');
    var f = await C.desfazerTudo();
    C.busy(btn, false); C.actualizarDesfazer();
    C.setStatus('desfazerStatus', f ? 'Algumas alterações não puderam ser desfeitas (' + f + ').' : 'As alterações do assistente foram desfeitas.', f ? 'err' : 'ok');
  }

  /* ============================================================ execução */
  function idDe(ids, n) { var id = ids[(+n || 0) - 1]; if (!id) throw new Error('o diapositivo ' + n + ' não existe'); return id; }
  function cor(c) { return /^#[0-9a-f]{6}$/i.test(c || '') ? c : null; }
  function caixa(posicao, w, h) {
    w = w || 560; h = h || 300;
    if (posicao === 'direita') return { left: LARG - w - 40, top: (ALT - h) / 2 + 20, width: w, height: h };
    if (posicao === 'baixo') return { left: (LARG - w) / 2, top: ALT - h - 30, width: w, height: h };
    if (posicao === 'topo') return { left: (LARG - w) / 2, top: 30, width: w, height: h };
    return { left: (LARG - w) / 2, top: (ALT - h) / 2 + 30, width: w, height: h };
  }
  async function idsFormas(ctx, slideId) { var sh = ctx.presentation.slides.getItem(slideId).shapes; sh.load('items/id'); await ctx.sync(); return sh.items.map(function (x) { return x.id; }); }
  function registarApagarNovas(slideId, antes) {
    C.registar('formas novas', function () {
      return PowerPoint.run(async function (ctx) {
        var sh = ctx.presentation.slides.getItem(slideId).shapes; sh.load('items/id'); await ctx.sync();
        sh.items.forEach(function (x) { if (antes.indexOf(x.id) < 0) x.delete(); });
        await ctx.sync();
      });
    });
  }
  async function inserirImagem(slideId, base64, pos) {
    if (!suporta('1.5')) throw new Error('esta versão do PowerPoint não permite inserir imagens num diapositivo escolhido; actualize o Microsoft 365');
    await PowerPoint.run(async function (ctx) {
      var antes = await idsFormas(ctx, slideId);
      if (suporta('1.5')) { ctx.presentation.setSelectedSlides([slideId]); await ctx.sync(); }
      await new Promise(function (ok, falha) {
        Office.context.document.setSelectedDataAsync(base64, { coercionType: Office.CoercionType.Image, imageLeft: pos.left, imageTop: pos.top, imageWidth: pos.width, imageHeight: pos.height },
          function (r) { if (r.status === Office.AsyncResultStatus.Failed) falha(new Error(r.error && r.error.message || 'não foi possível inserir a imagem')); else ok(); });
      });
      registarApagarNovas(slideId, antes);
    });
  }

  async function executar(a, ids, slides) {
    switch (a.tipo) {
      case 'texto': {
        var sid = idDe(ids, a.diapositivo), novo = String(a.texto || '').replace(/\\n/g, '\n');
        if (!novo.trim()) throw new Error('texto vazio');
        return PowerPoint.run(async function (ctx) {
          var f = ctx.presentation.slides.getItem(sid).shapes.getItem(String(a.forma));
          var t = f.textFrame.textRange; t.load('text'); await ctx.sync();
          var antigo = t.text;
          t.text = novo; await ctx.sync();
          C.registar('texto', function () { return PowerPoint.run(async function (c2) { c2.presentation.slides.getItem(sid).shapes.getItem(String(a.forma)).textFrame.textRange.text = antigo; await c2.sync(); }); });
        });
      }
      case 'novoDiapositivo': {
        return PowerPoint.run(async function (ctx) {
          var mestres = ctx.presentation.slideMasters; mestres.load('items/id');
          await ctx.sync();
          var mestre = mestres.items[0], lay = mestre.layouts; lay.load('items/id,items/name'); await ctx.sync();
          var escolhido = lay.items.find(function (l) { return /title and content|t[ií]tulo e conte[uú]do|t[ií]tulo y objetos|t[ií]tulo e objeto/i.test(l.name); }) || lay.items[1] || lay.items[0];
          ctx.presentation.slides.add({ slideMasterId: mestre.id, layoutId: escolhido.id });
          var sl = ctx.presentation.slides; sl.load('items/id'); await ctx.sync();
          var nova = sl.items[sl.items.length - 1], nid = nova.id;
          if (a.depoisDe && +a.depoisDe < sl.items.length - 1 && suporta('1.8')) { nova.moveTo(+a.depoisDe); await ctx.sync(); }
          var sh = nova.shapes; sh.load('items/id,items/top'); await ctx.sync();
          var formas = sh.items.slice().sort(function (x, y) { return x.top - y.top; });
          var titulo = String(a.titulo || ''), corpo = (a.pontos || []).map(String).join('\n');
          if (formas.length >= 2) { formas[0].textFrame.textRange.text = titulo; formas[1].textFrame.textRange.text = corpo; }
          else {
            nova.shapes.addTextBox(titulo, { left: 60, top: 40, width: LARG - 120, height: 70 });
            nova.shapes.addTextBox(corpo, { left: 60, top: 130, width: LARG - 120, height: ALT - 180 });
          }
          await ctx.sync();
          C.registar('diapositivo', function () { return PowerPoint.run(async function (c2) { c2.presentation.slides.getItem(nid).delete(); await c2.sync(); }); });
        });
      }
      case 'estilo': {
        var lista = Array.isArray(a.diapositivos) ? a.diapositivos.map(function (n) { return idDe(ids, n); }) : ids;
        return PowerPoint.run(async function (ctx) {
          var alvos = [];
          slides.filter(function (s) { return lista.indexOf(s.id) >= 0; }).forEach(function (s) {
            s.formas.forEach(function (f) {
              if (!f.texto.trim()) return;
              var t = ctx.presentation.slides.getItem(s.id).shapes.getItem(f.id).textFrame.textRange;
              t.font.load('name,size,color');
              alvos.push({ sid: s.id, fid: f.id, titulo: eTitulo(f, s.formas), t: t });
            });
          });
          await ctx.sync();
          var antes = alvos.map(function (x) { return { sid: x.sid, fid: x.fid, name: x.t.font.name, size: x.t.font.size, color: x.t.font.color }; });
          alvos.forEach(function (x) {
            if (a.fonte) x.t.font.name = String(a.fonte);
            var tam = x.titulo ? a.tamanhoTitulos : a.tamanhoTexto; if (tam) x.t.font.size = Math.max(8, Math.min(60, +tam));
            var c = cor(x.titulo ? a.corTitulos : a.corTexto); if (c) x.t.font.color = c;
          });
          await ctx.sync();
          C.registar('estilo', function () {
            return PowerPoint.run(async function (c2) {
              antes.forEach(function (o) { var f = c2.presentation.slides.getItem(o.sid).shapes.getItem(o.fid).textFrame.textRange.font; if (o.name) f.name = o.name; if (o.size) f.size = o.size; if (o.color) f.color = o.color; });
              await c2.sync();
            });
          });
        });
      }
      case 'grafico': {
        var sg = idDe(ids, a.diapositivo), pos = caixa(a.posicao, 520, 320);
        var png = desenharGrafico(a, pos.width * 2, pos.height * 2);
        return inserirImagem(sg, png, pos);
      }
      case 'tabela': {
        var stb = idDe(ids, a.diapositivo), linhas = (a.linhas || []).filter(Array.isArray).slice(0, 15);
        if (!linhas.length) throw new Error('tabela vazia');
        var nc = Math.min(8, Math.max.apply(null, linhas.map(function (l) { return l.length; })));
        linhas = linhas.map(function (l) { var x = l.slice(0, nc).map(function (v) { return v == null ? '' : String(v); }); while (x.length < nc) x.push(''); return x; });
        var pt = caixa(a.posicao, Math.min(840, 140 * nc), Math.min(380, 34 * linhas.length + 10));
        if (suporta('1.8')) {
          return PowerPoint.run(async function (ctx) {
            var antes = await idsFormas(ctx, stb);
            ctx.presentation.slides.getItem(stb).shapes.addTable(linhas.length, nc, { values: linhas, left: pt.left, top: pt.top, width: pt.width });
            await ctx.sync();
            registarApagarNovas(stb, antes);
          });
        }
        return inserirImagem(stb, desenharTabela(linhas, pt.width * 2, pt.height * 2), pt);
      }
      case 'caixaTexto': {
        var sc = idDe(ids, a.diapositivo), texto = String(a.texto || '').replace(/\\n/g, '\n');
        if (!texto.trim()) throw new Error('texto vazio');
        var pc = caixa(a.posicao, 760, a.posicao === 'baixo' || a.posicao === 'topo' ? 60 : 200);
        return PowerPoint.run(async function (ctx) {
          var antes = await idsFormas(ctx, sc);
          var tb = ctx.presentation.slides.getItem(sc).shapes.addTextBox(texto, pc);
          tb.textFrame.textRange.font.size = 16; tb.textFrame.textRange.font.color = '#3A3530';
          await ctx.sync();
          registarApagarNovas(sc, antes);
        });
      }
    }
    throw new Error('acção desconhecida');
  }

  /* ======================================================= gráficos e tabelas */
  var PALETA = ['#1E1A18', '#B08D57', '#7A5C3A', '#8A2E26', '#5C544F', '#D6C2A0'];
  function tela(w, h) { var c = document.createElement('canvas'); c.width = w; c.height = h; var g = c.getContext('2d'); g.fillStyle = '#FFFFFF'; g.fillRect(0, 0, w, h); return { c: c, g: g }; }
  function fmtNum(v) { return Number(v).toLocaleString('pt-PT', { maximumFractionDigits: 2 }); }
  function desenharGrafico(a, W, H) {
    var T = tela(W, H), g = T.g, s = W / 1040;
    var rot = (a.rotulos || []).map(String), series = (a.series || []).filter(function (x) { return x && Array.isArray(x.valores); }).slice(0, 6);
    if (!series.length) throw new Error('gráfico sem dados');
    var n = Math.max(rot.length, Math.max.apply(null, series.map(function (x) { return x.valores.length; })));
    var fonte = function (px, peso) { g.font = (peso || 400) + ' ' + Math.round(px * s) + 'px Calibri, Arial, sans-serif'; };
    var topo = 30 * s;
    if (a.titulo) { fonte(34, 700); g.fillStyle = '#1E1A18'; g.textAlign = 'left'; g.fillText(String(a.titulo), 40 * s, 60 * s); topo = 95 * s; }
    if (a.grafico === 'circular') {
      var vals = series[0].valores.map(Number), tot = vals.reduce(function (x, y) { return x + (y > 0 ? y : 0); }, 0) || 1;
      var cx = W * 0.32, cy = topo + (H - topo) / 2, r = Math.min(W * 0.26, (H - topo) / 2 - 20 * s), ang = -Math.PI / 2;
      vals.forEach(function (v, i) { var d = Math.max(0, v) / tot * Math.PI * 2; g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, r, ang, ang + d); g.closePath(); g.fillStyle = PALETA[i % PALETA.length]; g.fill(); g.strokeStyle = '#fff'; g.lineWidth = 3 * s; g.stroke(); ang += d; });
      fonte(24); g.textAlign = 'left';
      vals.forEach(function (v, i) { var y = topo + 40 * s + i * 42 * s; g.fillStyle = PALETA[i % PALETA.length]; g.fillRect(W * 0.62, y - 18 * s, 22 * s, 22 * s); g.fillStyle = '#3A3530'; g.fillText((rot[i] || '') + ': ' + fmtNum(v) + ' (' + Math.round(Math.max(0, v) / tot * 100) + ' %)', W * 0.62 + 34 * s, y); });
      return T.c.toDataURL('image/png').split(',')[1];
    }
    var todos = [].concat.apply([], series.map(function (x) { return x.valores.map(Number); }));
    var max = Math.max.apply(null, todos.concat([0])), min = Math.min.apply(null, todos.concat([0]));
    var esc = escala(min, max);
    var esq = 110 * s, dir = W - 40 * s, base = H - (series.length > 1 ? 110 : 70) * s, alto = topo + 20 * s;
    var horiz = a.grafico === 'barras';
    fonte(20); g.fillStyle = '#5C544F'; g.strokeStyle = '#E6DFD3'; g.lineWidth = 2 * s;
    esc.marcas.forEach(function (m) {
      if (horiz) { var x = esq + (m - esc.min) / (esc.max - esc.min) * (dir - esq); g.beginPath(); g.moveTo(x, alto); g.lineTo(x, base); g.stroke(); g.textAlign = 'center'; g.fillText(fmtNum(m), x, base + 30 * s); }
      else { var y = base - (m - esc.min) / (esc.max - esc.min) * (base - alto); g.beginPath(); g.moveTo(esq, y); g.lineTo(dir, y); g.stroke(); g.textAlign = 'right'; g.fillText(fmtNum(m), esq - 12 * s, y + 7 * s); }
    });
    var ns = series.length;
    if (a.grafico === 'linhas') {
      var passo = (dir - esq) / Math.max(1, n - 1);
      series.forEach(function (se, k) {
        g.strokeStyle = PALETA[k % PALETA.length]; g.fillStyle = g.strokeStyle; g.lineWidth = 5 * s; g.beginPath();
        se.valores.forEach(function (v, i) { var x = esq + i * passo, y = base - (v - esc.min) / (esc.max - esc.min) * (base - alto); if (i) g.lineTo(x, y); else g.moveTo(x, y); });
        g.stroke();
        se.valores.forEach(function (v, i) { var x = esq + i * passo, y = base - (v - esc.min) / (esc.max - esc.min) * (base - alto); g.beginPath(); g.arc(x, y, 7 * s, 0, Math.PI * 2); g.fill(); });
      });
      fonte(20); g.fillStyle = '#3A3530'; g.textAlign = 'center';
      rot.forEach(function (r, i) { g.fillText(r, esq + i * passo, base + 32 * s); });
    } else if (horiz) {
      var bandaH = (base - alto) / n, bh = bandaH * 0.7 / ns;
      for (var i = 0; i < n; i++) {
        series.forEach(function (se, k) { var v = +se.valores[i] || 0, x0 = esq + (0 - esc.min) / (esc.max - esc.min) * (dir - esq), x1 = esq + (v - esc.min) / (esc.max - esc.min) * (dir - esq); g.fillStyle = PALETA[k % PALETA.length]; g.fillRect(Math.min(x0, x1), alto + i * bandaH + bandaH * 0.15 + k * bh, Math.abs(x1 - x0), bh * 0.92); });
        fonte(20); g.fillStyle = '#3A3530'; g.textAlign = 'right'; g.fillText(rot[i] || '', esq - 12 * s, alto + i * bandaH + bandaH / 2 + 7 * s);
      }
    } else {
      var banda = (dir - esq) / n, bw = banda * 0.7 / ns;
      for (var j = 0; j < n; j++) {
        series.forEach(function (se, k) { var v = +se.valores[j] || 0, y0 = base - (0 - esc.min) / (esc.max - esc.min) * (base - alto), y1 = base - (v - esc.min) / (esc.max - esc.min) * (base - alto); g.fillStyle = PALETA[k % PALETA.length]; g.fillRect(esq + j * banda + banda * 0.15 + k * bw, Math.min(y0, y1), bw * 0.92, Math.abs(y1 - y0)); if (ns === 1) { fonte(19, 600); g.fillStyle = '#1E1A18'; g.textAlign = 'center'; g.fillText(fmtNum(v), esq + j * banda + banda / 2, Math.min(y0, y1) - 10 * s); } });
        fonte(20); g.fillStyle = '#3A3530'; g.textAlign = 'center'; g.fillText(rot[j] || '', esq + j * banda + banda / 2, base + 32 * s);
      }
    }
    if (ns > 1) {
      fonte(20); g.textAlign = 'left'; var lx = esq;
      series.forEach(function (se, k) { g.fillStyle = PALETA[k % PALETA.length]; g.fillRect(lx, H - 50 * s, 20 * s, 20 * s); g.fillStyle = '#3A3530'; var t = String(se.nome || 'Série ' + (k + 1)); g.fillText(t, lx + 30 * s, H - 33 * s); lx += 60 * s + g.measureText(t).width; });
    }
    return T.c.toDataURL('image/png').split(',')[1];
  }
  function escala(min, max) {
    if (max === min) max = min + 1;
    var bruto = (max - min) / 5, mag = Math.pow(10, Math.floor(Math.log10(bruto))), passo = [1, 2, 2.5, 5, 10].map(function (k) { return k * mag; }).find(function (p) { return p >= bruto; });
    var lo = Math.floor(min / passo) * passo, hi = Math.ceil(max / passo) * passo, marcas = [];
    for (var v = lo; v <= hi + passo / 2; v += passo) marcas.push(Math.round(v * 1e6) / 1e6);
    return { min: lo, max: hi, marcas: marcas };
  }
  function desenharTabela(linhas, W, H) {
    var T = tela(W, H), g = T.g, nl = linhas.length, nc = linhas[0].length, lh = H / nl, cw = W / nc;
    linhas.forEach(function (l, i) {
      g.fillStyle = i === 0 ? '#1E1A18' : (i % 2 ? '#FFFFFF' : '#F4F1EC'); g.fillRect(0, i * lh, W, lh);
      l.forEach(function (v, j) { g.fillStyle = i === 0 ? '#FFFFFF' : '#1E1A18'; g.font = (i === 0 ? 700 : 400) + ' ' + Math.round(lh * 0.42) + 'px Calibri, Arial, sans-serif'; g.textAlign = 'left'; g.fillText(String(v).slice(0, 60), j * cw + 14, i * lh + lh * 0.64, cw - 24); });
    });
    g.strokeStyle = '#D6CEBF'; g.lineWidth = 2; g.strokeRect(1, 1, W - 2, H - 2);
    return T.c.toDataURL('image/png').split(',')[1];
  }

  /* ============================================================ anonimizar */
  async function analisar() {
    if (!C.licenciado()) return;
    var btn = $('btnAnalyze'); C.busy(btn, true, 'A analisar…');
    C.setStatus('analyzeStatus', 'A ler a apresentação…', 'busy');
    try {
      var slides = await lerDiapositivos();
      st.textos = [];
      slides.forEach(function (s) { s.formas.forEach(function (f) { if (f.texto.trim()) st.textos.push({ sid: s.id, n: s.n, fid: f.id, t: f.texto }); }); });
      await C.tick();
      var txt = st.textos.map(function (x) { return x.t; });
      var existentes = E.existingLabels(txt);
      st.applied.forEach(function (en) { var mm = en.label.match(/^\[([A-Z_]+?)_(\d+)\]$/); if (mm) existentes[mm[1]] = Math.max(existentes[mm[1]] || 0, +mm[2]); });
      st.entities = E.analyze(txt, { types: C.tiposActivos(), existingLabels: existentes }).entities;
      C.mostrarEntidades(st.entities, localizar);
      $('review').hidden = false; $('applyCard').hidden = false;
      C.setStatus('analyzeStatus', 'Análise concluída: ' + C.plural(slides.length, 'diapositivo lido.', 'diapositivos lidos.'), 'ok');
    } catch (e) { C.setStatus('analyzeStatus', C.errMsg(e), 'err'); }
    finally { C.busy(btn, false); }
  }
  function acrescentarTermo() {
    var term = $('manualTerm').value.trim(); if (!term) return;
    var en = E.addManual(st.entities.concat(st.applied), term, $('manualType').value);
    if (en && st.entities.indexOf(en) < 0 && st.applied.indexOf(en) < 0) st.entities.push(en);
    $('manualTerm').value = '';
    C.mostrarEntidades(st.entities, localizar);
  }
  async function localizar(en) {
    if (!suporta('1.5')) return;
    var vars = Array.from(en.variants.keys());
    var x = st.textos.find(function (t) { return vars.some(function (v) { return t.t.indexOf(v) >= 0; }); });
    if (!x) return;
    await PowerPoint.run(async function (ctx) { ctx.presentation.setSelectedSlides([x.sid]); await ctx.sync(); });
  }
  // aplica as substituições preservando a formatação, quando a versão do PowerPoint o permite
  function operacoes(texto, plano) {
    var ops = [], atual = texto;
    plano.forEach(function (nivel) {
      nivel.forEach(function (it) {
        var idx = E.wordSearch(atual, it.search, it.matchCase, it.wholeWord);
        for (var k = idx.length - 1; k >= 0; k--) {
          ops.push({ ini: idx[k], len: it.search.length, novo: it.replacement });
          atual = atual.slice(0, idx[k]) + it.replacement + atual.slice(idx[k] + it.search.length);
        }
      });
    });
    return { ops: ops, final: atual };
  }
  async function substituir(itens, transformar) {
    var sub = suporta('1.5'), alteradas = 0;
    await PowerPoint.run(async function (ctx) {
      for (var i = 0; i < itens.length; i++) {
        var x = itens[i], r = transformar(x.t);
        if (r.final === x.t) continue;
        var tr = ctx.presentation.slides.getItem(x.sid).shapes.getItem(x.fid).textFrame.textRange;
        if (sub && r.ops) r.ops.forEach(function (o) { tr.getSubstring(o.ini, o.len).text = o.novo; });
        else tr.text = r.final;
        alteradas++;
        if (alteradas % 100 === 0) await ctx.sync();
      }
      await ctx.sync();
    });
    return alteradas;
  }
  async function anonimizar() {
    if (!C.licenciado()) return;
    var sel = st.entities.filter(function (e) { return e.selected; });
    if (!sel.length) { C.setStatus('applyStatus', 'Não há elementos seleccionados.', 'err'); return; }
    var modo = (document.querySelector('input[name=mode]:checked') || {}).value || 'pseudo';
    var btn = $('btnApply'); C.busy(btn, true, 'A aplicar…');
    try {
      var plano = E.buildPlan(sel, { style: modo === 'x' ? 'x' : 'label' });
      var n = await substituir(st.textos, function (t) { return operacoes(t, plano); });
      st.applied = st.applied.concat(sel);
      $('resultSummary').innerHTML = '<p>' + C.plural(n, 'caixa de texto alterada.', 'caixas de texto alteradas.') + '</p>';
      $('tableBlock').hidden = modo !== 'pseudo';
      if (modo === 'pseudo') $('tableText').value = E.buildTable(st.applied, { document: st.nome || 'apresentação' });
      $('resultCard').hidden = false;
      C.setStatus('applyStatus', 'Substituição concluída.', 'ok');
      st.entities = st.entities.filter(function (e) { return !e.selected; });
      C.mostrarEntidades(st.entities, localizar);
      $('resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) { C.setStatus('applyStatus', C.errMsg(e), 'err'); }
    finally { C.busy(btn, false); }
  }
  function copiarTabela() {
    var ta = $('tableText'); ta.select();
    var ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    if (!ok && navigator.clipboard) navigator.clipboard.writeText(ta.value).then(function () { C.setStatus('tableStatus', 'Tabela copiada.', 'ok'); });
    else C.setStatus('tableStatus', ok ? 'Tabela copiada.' : 'Seleccione o texto e copie-o manualmente.', ok ? 'ok' : 'err');
  }
  async function repor() {
    if (!C.licenciado()) return;
    var mapa = E.parseTable($('revertText').value);
    if (!mapa.length) { C.setStatus('revertStatus', 'A tabela não contém correspondências válidas.', 'err'); return; }
    var btn = $('btnRevert'); C.busy(btn, true, 'A repor…');
    try {
      var slides = await lerDiapositivos(), itens = [];
      slides.forEach(function (s) { s.formas.forEach(function (f) { if (f.texto.trim()) itens.push({ sid: s.id, fid: f.id, t: f.texto }); }); });
      var n = await substituir(itens, function (t) {
        // repõe etiqueta a etiqueta, para manter a formatação
        var ops = [], atual = t;
        mapa.forEach(function (r) { var idx = E.wordSearch(atual, r.label, true, false); for (var k = idx.length - 1; k >= 0; k--) { ops.push({ ini: idx[k], len: r.label.length, novo: r.original }); atual = atual.slice(0, idx[k]) + r.original + atual.slice(idx[k] + r.label.length); } });
        return { ops: ops, final: atual };
      });
      C.setStatus('revertStatus', 'Dados repostos em ' + C.plural(n, 'caixa de texto.', 'caixas de texto.'), 'ok');
    } catch (e) { C.setStatus('revertStatus', C.errMsg(e), 'err'); }
    finally { C.busy(btn, false); }
  }

  window.__cofrePpt = { st: st, desenharGrafico: desenharGrafico };
  boot();
})();
