/*
 * Assessor para Excel: painel lateral.
 * Assistente com instruções (plano de acções executado com a API do Excel, com desfazer)
 * e anonimização ou pseudonimização das células com texto, feita no computador do utilizador.
 * Ana Marta Soares, Especialista em IA aplicada ao Direito.
 */
'use strict';
(function () {
  var C = window.Cofre, E = window.AnonEngine, $ = C.$;
  var st = { plano: null, mascara: null, cells: [], entities: [], applied: [], livro: '', api19: false };

  /* ============================================================ arranque */
  function boot() {
    if (!window.__cofreSuportado || !E || !C) { $('unsupported').hidden = false; return; }
    C.iniciar({ produto: 'excel', nome: 'Assessor para Excel' });
    bind();
    if (typeof Office === 'undefined') { $('outside').hidden = false; return; }
    Office.onReady(function (info) {
      if (!info || info.host !== Office.HostType.Excel) { $('outside').hidden = false; return; }
      if (!suporta('1.7')) { $('unsupported').hidden = false; return; }
      st.api19 = suporta('1.9');
      try { var u = Office.context.document && Office.context.document.url; if (u) st.livro = decodeURIComponent(String(u).split(/[\\/]/).pop()); } catch (e) { /* sem nome */ }
    });
  }
  function suporta(v) { try { return Office.context.requirements.isSetSupported('ExcelApi', v); } catch (e) { return false; } }
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

  /* ======================================================= endereços e fórmulas */
  function colNum(l) { var n = 0; for (var i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64); return n; }
  function colLetra(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function semFolha(a) { return String(a || '').replace(/^.*!/, '').replace(/\$/g, '').trim().toUpperCase(); }
  var RE_CEL = /^([A-Z]{1,3})(\d+)$/;
  function endereco(a) {
    a = semFolha(a);
    if (/^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$/.test(a) || /^[A-Z]{1,3}:[A-Z]{1,3}$/.test(a) || /^\d+:\d+$/.test(a)) return a;
    throw new Error('Endereço inválido: ' + a);
  }
  function partes(a) { var p = endereco(a).split(':'); var i = p[0].match(RE_CEL), f = (p[1] || p[0]).match(RE_CEL); if (!i || !f) return null; return { c1: colNum(i[1]), r1: +i[2], c2: colNum(f[1]), r2: +f[2] }; }
  // desloca as referências relativas de uma fórmula (sem tocar no texto entre aspas nem nos nomes de funções)
  function deslocarFormula(f, dr, dc) {
    var out = '', dentro = false;
    var re = /(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g;
    String(f).split('"').forEach(function (seg, i) {
      if (i % 2 === 1) { out += '"' + seg + '"'; return; }
      out += (i > 0 ? '' : '') + seg.replace(re, function (m, ac, col, ar, row, pos, s) {
        var antes = s[pos - 1];
        if (antes && /[A-Za-z_.]/.test(antes)) return m;
        var c = ac ? col : colLetra(colNum(col) + dc);
        var r = ar ? row : String(+row + dr);
        return ac + c + ar + r;
      });
    });
    return out;
  }

  /* ============================================================ contexto */
  async function lerContexto(ambito) {
    return Excel.run(async function (ctx) {
      var wb = ctx.workbook, ws = wb.worksheets.getActiveWorksheet();
      ws.load('name');
      var sel = wb.getSelectedRange(); sel.load('address,cellCount');
      var folhas = wb.worksheets; folhas.load('items/name');
      var graf = ws.charts; graf.load('items/name');
      var tabs = ws.tables; tabs.load('items/name');
      await ctx.sync();
      var alvo = (ambito === 'folha' || sel.cellCount <= 1) ? ws.getUsedRangeOrNullObject(true) : sel;
      alvo.load('address,values,numberFormat,rowCount,columnCount,rowIndex,columnIndex');
      await ctx.sync();
      if (alvo.isNullObject) return { folha: ws.name, vazio: true, folhas: folhas.items.map(function (f) { return f.name; }) };
      return {
        folha: ws.name, folhas: folhas.items.map(function (f) { return f.name; }),
        graficos: graf.items.map(function (g) { return g.name; }), tabelas: tabs.items.map(function (t) { return t.name; }),
        endereco: semFolha(alvo.address), linhas: alvo.rowCount, colunas: alvo.columnCount,
        r0: alvo.rowIndex + 1, c0: alvo.columnIndex + 1, valores: alvo.values, formatos: alvo.numberFormat
      };
    });
  }
  function textoContexto(cx, m) {
    if (cx.vazio) return 'FOLHA ACTIVA: ' + cx.folha + ' (vazia)\nFOLHAS DO LIVRO: ' + cx.folhas.join(', ');
    var MAXL = 150, MAXC = 30, linhas = [];
    var nl = cx.valores.length, nc = Math.min(cx.valores[0].length, MAXC);
    var idx = [];
    for (var i = 0; i < nl; i++) idx.push(i);
    if (nl > MAXL) idx = idx.slice(0, 120).concat(idx.slice(nl - 10));
    idx.forEach(function (i, k) {
      if (k > 0 && i !== idx[k - 1] + 1) linhas.push('(… ' + (i - idx[k - 1] - 1) + ' linhas omitidas …)');
      var cel = [];
      for (var j = 0; j < nc; j++) {
        var v = cx.valores[i][j];
        if (v === '' || v == null) continue;
        if (typeof v === 'string') v = m.tapar(v);
        else if (grandeInteiro(v)) { var sv = String(v), tv = m.tapar(sv); if (tv !== sv) v = tv; }
        var fmt = cx.formatos[i] && cx.formatos[i][j];
        cel.push(colLetra(cx.c0 + j) + (cx.r0 + i) + '=' + JSON.stringify(v) + (fmt && fmt !== 'General' && typeof v === 'number' ? ' [' + fmt + ']' : ''));
      }
      if (cel.length) linhas.push(cel.join('; '));
    });
    // estatísticas das colunas numéricas, calculadas sobre todas as linhas
    var est = [];
    for (var c = 0; c < nc; c++) {
      var nums = [];
      for (var r = 0; r < nl; r++) if (typeof cx.valores[r][c] === 'number') nums.push(cx.valores[r][c]);
      if (nums.length >= 3) {
        var soma = nums.reduce(function (a, b) { return a + b; }, 0);
        est.push(colLetra(cx.c0 + c) + ': ' + nums.length + ' números, soma ' + arred(soma) + ', média ' + arred(soma / nums.length) + ', mínimo ' + Math.min.apply(null, nums) + ', máximo ' + Math.max.apply(null, nums));
      }
    }
    return 'FOLHA ACTIVA: ' + m.tapar(cx.folha) + '\nFOLHAS DO LIVRO: ' + cx.folhas.map(m.tapar).join(', ') +
      '\nGRÁFICOS NA FOLHA: ' + (cx.graficos.join(', ') || 'nenhum') + '\nTABELAS NA FOLHA: ' + (cx.tabelas.join(', ') || 'nenhuma') +
      '\nINTERVALO DOS DADOS: ' + cx.endereco + ' (' + cx.linhas + ' linhas, ' + cx.colunas + ' colunas; a primeira linha é ' + cx.r0 + ', a última é ' + (cx.r0 + nl - 1) + ')' +
      '\nCÉLULAS PREENCHIDAS (célula=valor):\n' + linhas.join('\n') + (est.length ? '\nESTATÍSTICAS DAS COLUNAS NUMÉRICAS:\n' + est.join('\n') : '');
  }
  // números inteiros longos (NIF, NISS, telefones guardados como números) são tratados como texto
  function grandeInteiro(v) { return typeof v === 'number' && Number.isInteger(v) && Math.abs(v) >= 1e8; }
  function arred(x) { return Math.round(x * 100) / 100; }

  var SISTEMA = [
    'És o Assessor para Excel, um assistente ao serviço de juristas, advogados e serviços públicos em Portugal. Recebes uma instrução e o conteúdo de uma folha de cálculo.',
    'Respondes SEMPRE e APENAS com um objecto JSON, sem texto antes ou depois, com esta forma:',
    '{"resposta":"texto curto em português europeu para o utilizador (o que vais fazer, ou a resposta se a instrução for uma pergunta)","accoes":[...]}',
    'Acções possíveis (usa só as necessárias, no máximo 12, pela ordem de execução):',
    '{"tipo":"grafico","intervalo":"A1:B13","grafico":"colunas|colunasEmpilhadas|barras|linhas|circular|area|dispersao|rosca","titulo":"..."}  (o intervalo inclui a linha de cabeçalhos e a coluna das categorias)',
    '{"tipo":"formula","celula":"F2","formula":"=SUM(B2:E2)","preencherAte":"F13","cabecalho":{"celula":"F1","texto":"Total"}}  (a fórmula é escrita para a primeira célula e copiada até preencherAte, com as referências relativas ajustadas)',
    '{"tipo":"valores","celula":"H1","valores":[["Indicador","Valor"],["Total","=SUM(F2:F13)"]]}  (escreve uma tabela de valores ou fórmulas a partir da célula indicada)',
    '{"tipo":"ordenar","intervalo":"A1:F13","coluna":"F","ascendente":false,"temCabecalho":true}',
    '{"tipo":"tabela","intervalo":"A1:F13","estilo":"TableStyleMedium2"}  (formata como tabela do Excel; não uses se o intervalo já pertencer a uma tabela)',
    '{"tipo":"formatoNumero","intervalo":"B2:F13","formato":"#,##0.00 €"}',
    '{"tipo":"realcar","intervalo":"F2:F13","regra":"acimaMedia|abaixoMedia|maiorQue|menorQue|igualA|contemTexto|escalaCores|barras","valor":1000,"cor":"#F4D9C6"}',
    '{"tipo":"formatar","intervalo":"A1:F1","negrito":true,"italico":false,"corFundo":"#1E1A18","corTexto":"#FFFFFF","alinhamento":"esquerda|centro|direita","tamanho":11,"contornos":true}',
    '{"tipo":"larguraAuto","intervalo":"A1:F13"}',
    '{"tipo":"filtro","intervalo":"A1:F13"}',
    '{"tipo":"congelar","linhas":1}',
    '{"tipo":"folha","nome":"Resumo"}  (cria uma folha; as acções seguintes podem indicar "folha":"Resumo" para escrever nela)',
    'Regras:',
    '1. As fórmulas usam os nomes das funções em inglês e a vírgula como separador de argumentos (por exemplo =SUMIF(A2:A13,"Lisboa",C2:C13)), porque são escritas através da API do Excel.',
    '2. Usa apenas endereços que existam nos dados fornecidos. Não inventes dados nem valores.',
    '3. As etiquetas como [PESSOA_1] ou [NIF_1] substituem dados pessoais: mantém-nas exactamente como estão.',
    '4. Escreve em português europeu (grafia anterior ao Acordo Ortográfico de 1990), em registo profissional, sem travessões.',
    '5. Se a instrução for apenas uma pergunta sobre os dados, responde em "resposta" e devolve "accoes":[].',
    '6. Se a instrução não puder ser cumprida com estas acções, explica em "resposta" e devolve "accoes":[].',
    '7. O intervalo de um gráfico tem de ser contíguo e conter só a coluna das categorias e as colunas numéricas a representar. Se essas colunas não estiverem lado a lado, cria primeiro, com a acção valores, uma pequena tabela auxiliar com fórmulas que remetam para as células de origem (por exemplo =C2 e =F2), e faz o gráfico sobre essa tabela.',
    '8. Para destacar as linhas ou as células de uma pessoa ou entidade, usa realcar com a regra contemTexto e o texto exacto.'
  ].join('\n');

  /* ============================================================ assistente */
  async function preparar() {
    if (!C.licenciado()) return;
    var instr = $('instrucao').value.trim();
    if (!instr) { C.setStatus('assistStatus', 'Escreva o que pretende.', 'err'); return; }
    var btn = $('btnExecutar');
    C.busy(btn, true, 'A preparar…');
    $('planoCard').hidden = true;
    try {
      C.setStatus('assistStatus', 'A ler a folha…', 'busy');
      var cx = await lerContexto($('ambito').value);
      var textos = [instr, cx.folha].concat((cx.folhas || []));
      (cx.valores || []).forEach(function (l) { l.forEach(function (v) { if (typeof v === 'string') textos.push(v); else if (grandeInteiro(v)) textos.push(String(v)); }); });
      var m = C.mascara(textos, $('optMascara').checked);
      st.mascara = m;
      C.setStatus('assistStatus', (m.n ? 'Dados pessoais substituídos por etiquetas antes do envio: ' + m.n + '. ' : '') + 'A preparar o plano…', 'busy');
      var txt = await C.pedirIA({ sistema: SISTEMA, mensagem: 'INSTRUÇÃO: ' + m.tapar(instr) + '\n\n' + textoContexto(cx, m), maxTokens: 4000, esforco: 'medio' });
      var j = C.destaparFundo(C.extrairJSON(txt), m);
      var accoes = (Array.isArray(j.accoes) ? j.accoes : []).filter(function (a) { return a && DESCR[a.tipo]; }).slice(0, 12);
      st.plano = { accoes: accoes, folha: cx.folha };
      mostrarPlano(j.resposta || '', accoes);
      C.setStatus('assistStatus', '');
    } catch (e) {
      C.setStatus('assistStatus', e.message || C.errMsg(e), 'err');
    } finally { C.busy(btn, false); }
  }
  var NOMES_GRAF = { colunas: 'colunas', colunasEmpilhadas: 'colunas empilhadas', barras: 'barras', linhas: 'linhas', circular: 'circular', area: 'área', dispersao: 'dispersão', rosca: 'anel' };
  var DESCR = {
    grafico: function (a) { return 'Gráfico de ' + (NOMES_GRAF[a.grafico] || 'colunas') + ' com os dados de ' + a.intervalo + (a.titulo ? ', com o título «' + a.titulo + '»' : '') + '.'; },
    formula: function (a) { return 'Fórmula ' + a.formula + ' em ' + a.celula + (a.preencherAte ? ', copiada até ' + a.preencherAte : '') + (a.cabecalho && a.cabecalho.texto ? ', com o cabeçalho «' + a.cabecalho.texto + '» em ' + a.cabecalho.celula : '') + '.'; },
    valores: function (a) { return 'Escrever ' + C.plural((a.valores || []).length, 'linha', 'linhas') + ' a partir de ' + a.celula + '.'; },
    ordenar: function (a) { return 'Ordenar ' + a.intervalo + ' pela coluna ' + a.coluna + ', ' + (a.ascendente === false ? 'do maior para o menor' : 'do menor para o maior') + '.'; },
    tabela: function (a) { return 'Formatar ' + a.intervalo + ' como tabela.'; },
    formatoNumero: function (a) { return 'Formato «' + a.formato + '» em ' + a.intervalo + '.'; },
    realcar: function (a) { return 'Realçar em ' + a.intervalo + ' ' + ({ acimaMedia: 'os valores acima da média', abaixoMedia: 'os valores abaixo da média', maiorQue: 'os valores maiores do que ' + a.valor, menorQue: 'os valores menores do que ' + a.valor, igualA: 'os valores iguais a ' + a.valor, contemTexto: 'as células que contêm «' + a.valor + '»', escalaCores: 'os valores com uma escala de cores', barras: 'os valores com barras de dados' }[a.regra] || 'os valores') + '.'; },
    formatar: function (a) { return 'Formatar ' + a.intervalo + [a.negrito ? 'negrito' : '', a.corFundo ? 'fundo ' + a.corFundo : '', a.corTexto ? 'texto ' + a.corTexto : '', a.alinhamento ? 'alinhamento ' + a.alinhamento : '', a.contornos ? 'contornos' : ''].filter(Boolean).map(function (x, i) { return (i ? ', ' : ': ') + x; }).join('') + '.'; },
    larguraAuto: function (a) { return 'Ajustar a largura das colunas de ' + a.intervalo + '.'; },
    filtro: function (a) { return 'Filtros nos cabeçalhos de ' + a.intervalo + '.'; },
    congelar: function (a) { return 'Fixar ' + C.plural(a.linhas || 1, 'linha', 'linhas') + ' no topo.'; },
    folha: function (a) { return 'Criar a folha «' + a.nome + '».'; }
  };
  function mostrarPlano(resp, accoes) {
    $('resposta').textContent = resp;
    var ol = $('plano'); ol.innerHTML = '';
    accoes.forEach(function (a) {
      var li = document.createElement('li');
      var d; try { d = DESCR[a.tipo](a); } catch (e) { d = a.tipo; }
      li.textContent = d + (a.folha ? ' (folha «' + a.folha + '»)' : '');
      ol.appendChild(li);
    });
    $('btnAplicar').hidden = !accoes.length;
    $('btnDescartar').textContent = accoes.length ? 'Descartar' : 'Fechar';
    C.setStatus('aplicarStatus', accoes.length ? '' : '');
    $('planoCard').hidden = false;
    $('planoCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function aplicarPlano() {
    if (!st.plano) return;
    var btn = $('btnAplicar'); C.busy(btn, true, 'A aplicar…');
    var itens = $('plano').children, ok = 0, falhas = [];
    for (var i = 0; i < st.plano.accoes.length; i++) {
      var a = st.plano.accoes[i];
      try { await executar(a); ok++; itens[i].classList.add('feito'); }
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
  function folhaDe(ctx, a) { return a.folha ? ctx.workbook.worksheets.getItem(a.folha) : ctx.workbook.worksheets.getActiveWorksheet(); }
  // guarda o conteúdo e o formato numérico de um intervalo, para desfazer
  async function fotografar(ctx, ws, end) {
    var r = ws.getRange(end); r.load('formulas,numberFormat'); await ctx.sync();
    var folha = ws; folha.load('name'); await ctx.sync();
    var nome = folha.name, f = r.formulas, n = r.numberFormat;
    return function () { return Excel.run(async function (c2) { var x = c2.workbook.worksheets.getItem(nome).getRange(end); x.formulas = f; x.numberFormat = n; await c2.sync(); }); };
  }
  async function fotografarFormato(ctx, ws, end) {
    if (!st.api19) return null;
    var r = ws.getRange(end);
    var props = r.getCellProperties({ format: { fill: { color: true }, font: { bold: true, italic: true, color: true, size: true }, horizontalAlignment: true } });
    ws.load('name'); await ctx.sync();
    var nome = ws.name, v = props.value;
    return function () { return Excel.run(async function (c2) { c2.workbook.worksheets.getItem(nome).getRange(end).setCellProperties(v); await c2.sync(); }); };
  }
  var TIPOS_GRAF = { colunas: 'ColumnClustered', colunasEmpilhadas: 'ColumnStacked', barras: 'BarClustered', linhas: 'LineMarkers', circular: 'Pie', area: 'Area', dispersao: 'XYScatter', rosca: 'Doughnut' };

  async function executar(a) {
    return Excel.run(async function (ctx) {
      var ws = folhaDe(ctx, a); ws.load('name'); await ctx.sync();
      var nomeFolha = ws.name;
      switch (a.tipo) {
        case 'grafico': {
          var end = endereco(a.intervalo), p = partes(end);
          var g = ws.charts.add(TIPOS_GRAF[a.grafico] || 'ColumnClustered', ws.getRange(end), 'Auto');
          if (a.titulo) { g.title.text = a.titulo; g.title.visible = true; }
          if (p) { var col = colLetra(p.c2 + 2); g.setPosition(col + p.r1, colLetra(p.c2 + 9) + (p.r1 + 15)); }
          g.load('name'); await ctx.sync();
          var nomeG = g.name;
          C.registar('gráfico', function () { return Excel.run(async function (c2) { c2.workbook.worksheets.getItem(nomeFolha).charts.getItem(nomeG).delete(); await c2.sync(); }); });
          return;
        }
        case 'formula': {
          var cel = endereco(a.celula), pc = partes(cel), fim = a.preencherAte ? partes(endereco(a.preencherAte)) : pc;
          var formula = String(a.formula || '').trim(); if (formula && formula[0] !== '=') formula = '=' + formula;
          if (!formula) throw new Error('fórmula vazia');
          var alvo = colLetra(Math.min(pc.c1, fim.c1)) + Math.min(pc.r1, fim.r1) + ':' + colLetra(Math.max(pc.c1, fim.c1)) + Math.max(pc.r1, fim.r1);
          if (a.cabecalho && a.cabecalho.celula) { var u0 = await fotografar(ctx, ws, endereco(a.cabecalho.celula)); C.registar('cabeçalho', u0); ws.getRange(endereco(a.cabecalho.celula)).values = [[String(a.cabecalho.texto || '')]]; }
          var u1 = await fotografar(ctx, ws, alvo); C.registar('fórmula', u1);
          var pa = partes(alvo), linhas = [];
          for (var r = pa.r1; r <= pa.r2; r++) { var l = []; for (var c = pa.c1; c <= pa.c2; c++) l.push(deslocarFormula(formula, r - pc.r1, c - pc.c1)); linhas.push(l); }
          ws.getRange(alvo).formulas = linhas;
          await ctx.sync();
          return;
        }
        case 'valores': {
          var v = (a.valores || []).filter(Array.isArray); if (!v.length) throw new Error('sem valores');
          var largura = Math.max.apply(null, v.map(function (l) { return l.length; }));
          v = v.map(function (l) { var x = l.slice(); while (x.length < largura) x.push(''); return x.map(function (y) { return y == null ? '' : y; }); });
          var pi = partes(endereco(a.celula));
          var area = colLetra(pi.c1) + pi.r1 + ':' + colLetra(pi.c1 + largura - 1) + (pi.r1 + v.length - 1);
          var u2 = await fotografar(ctx, ws, area); C.registar('valores', u2);
          ws.getRange(area).formulas = v;
          await ctx.sync();
          return;
        }
        case 'ordenar': {
          var endO = endereco(a.intervalo), po = partes(endO);
          var u3 = await fotografar(ctx, ws, endO); C.registar('ordenação', u3);
          var k = /^[A-Z]{1,3}$/.test(String(a.coluna || '').toUpperCase()) ? colNum(String(a.coluna).toUpperCase()) - po.c1 : (+a.coluna || 0);
          if (k < 0 || k > po.c2 - po.c1) throw new Error('coluna de ordenação fora do intervalo');
          ws.getRange(endO).sort.apply([{ key: k, ascending: a.ascendente !== false }], false, a.temCabecalho !== false);
          await ctx.sync();
          return;
        }
        case 'tabela': {
          var endT = endereco(a.intervalo);
          var uf = await fotografarFormato(ctx, ws, endT);
          var t = ws.tables.add(endT, true);
          t.style = /^TableStyle(Light|Medium|Dark)\d+$/.test(a.estilo || '') ? a.estilo : 'TableStyleMedium2';
          t.load('name'); await ctx.sync();
          var nomeT = t.name;
          C.registar('tabela', function () { return Excel.run(async function (c2) { c2.workbook.worksheets.getItem(nomeFolha).tables.getItem(nomeT).convertToRange(); await c2.sync(); if (uf) await uf(); }); });
          return;
        }
        case 'formatoNumero': {
          var endN = endereco(a.intervalo);
          var u4 = await fotografar(ctx, ws, endN); C.registar('formato', u4);
          var rn = ws.getRange(endN); rn.load('rowCount,columnCount'); await ctx.sync();
          var fmt = String(a.formato || 'General'), mat = [];
          for (var i = 0; i < rn.rowCount; i++) { var lf = []; for (var j = 0; j < rn.columnCount; j++) lf.push(fmt); mat.push(lf); }
          rn.numberFormat = mat; await ctx.sync();
          return;
        }
        case 'realcar': {
          var endR = endereco(a.intervalo), rr = ws.getRange(endR), cor = /^#[0-9a-f]{6}$/i.test(a.cor || '') ? a.cor : '#F4D9C6', cf;
          var regra = a.regra;
          if (regra === 'acimaMedia' || regra === 'abaixoMedia') {
            cf = rr.conditionalFormats.add('PresetCriteria');
            cf.preset.format.fill.color = cor;
            cf.preset.rule = { criterion: regra === 'acimaMedia' ? 'AboveAverage' : 'BelowAverage' };
          } else if (regra === 'igualA' && (a.valor === '' || isNaN(Number(a.valor)))) {
            // igualdade a um texto (por exemplo, o nome de um cliente)
            cf = rr.conditionalFormats.add('ContainsText');
            cf.textComparison.format.fill.color = cor;
            cf.textComparison.rule = { operator: 'Contains', text: String(a.valor || '') };
          } else if (regra === 'maiorQue' || regra === 'menorQue' || regra === 'igualA') {
            cf = rr.conditionalFormats.add('CellValue');
            cf.cellValue.format.fill.color = cor;
            cf.cellValue.rule = { formula1: String(Number(a.valor) || 0), operator: { maiorQue: 'GreaterThan', menorQue: 'LessThan', igualA: 'EqualTo' }[regra] };
          } else if (regra === 'contemTexto') {
            cf = rr.conditionalFormats.add('ContainsText');
            cf.textComparison.format.fill.color = cor;
            cf.textComparison.rule = { operator: 'Contains', text: String(a.valor || '') };
          } else if (regra === 'barras') {
            cf = rr.conditionalFormats.add('DataBar');
            cf.dataBar.positiveFormat.fillColor = cor;
          } else {
            cf = rr.conditionalFormats.add('ColorScale');
            cf.colorScale.criteria = { minimum: { formula: null, type: 'LowestValue', color: '#FFFFFF' }, maximum: { formula: null, type: 'HighestValue', color: cor } };
          }
          cf.load('id'); await ctx.sync();
          var idCf = cf.id;
          C.registar('realce', function () { return Excel.run(async function (c2) { c2.workbook.worksheets.getItem(nomeFolha).getRange(endR).conditionalFormats.getItem(idCf).delete(); await c2.sync(); }); });
          return;
        }
        case 'formatar': {
          var endF = endereco(a.intervalo), rf = ws.getRange(endF);
          var u5 = await fotografarFormato(ctx, ws, endF); if (u5) C.registar('formatação', u5);
          if (a.negrito != null) rf.format.font.bold = !!a.negrito;
          if (a.italico != null) rf.format.font.italic = !!a.italico;
          if (/^#[0-9a-f]{6}$/i.test(a.corFundo || '')) rf.format.fill.color = a.corFundo;
          if (/^#[0-9a-f]{6}$/i.test(a.corTexto || '')) rf.format.font.color = a.corTexto;
          if (a.tamanho) rf.format.font.size = Math.max(6, Math.min(48, +a.tamanho));
          var al = { esquerda: 'Left', centro: 'Center', direita: 'Right' }[a.alinhamento]; if (al) rf.format.horizontalAlignment = al;
          if (a.contornos) ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'].forEach(function (b) { var x = rf.format.borders.getItem(b); x.style = 'Continuous'; x.color = '#BFB6A8'; });
          await ctx.sync();
          return;
        }
        case 'larguraAuto': { ws.getRange(endereco(a.intervalo)).format.autofitColumns(); await ctx.sync(); return; }
        case 'filtro': {
          ws.autoFilter.apply(ws.getRange(endereco(a.intervalo))); await ctx.sync();
          C.registar('filtro', function () { return Excel.run(async function (c2) { c2.workbook.worksheets.getItem(nomeFolha).autoFilter.remove(); await c2.sync(); }); });
          return;
        }
        case 'congelar': {
          ws.freezePanes.freezeRows(Math.max(1, Math.min(10, +a.linhas || 1))); await ctx.sync();
          C.registar('fixar', function () { return Excel.run(async function (c2) { c2.workbook.worksheets.getItem(nomeFolha).freezePanes.unfreeze(); await c2.sync(); }); });
          return;
        }
        case 'folha': {
          var nome = String(a.nome || 'Resumo').replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31).trim() || 'Resumo';
          var nova = ctx.workbook.worksheets.add(nome); nova.load('name'); await ctx.sync();
          var nn = nova.name;
          C.registar('folha', function () { return Excel.run(async function (c2) { c2.workbook.worksheets.getItem(nn).delete(); await c2.sync(); }); });
          return;
        }
      }
      throw new Error('acção desconhecida');
    });
  }

  /* ============================================================ anonimizar */
  async function lerCelulas() {
    return Excel.run(async function (ctx) {
      var folhas = ctx.workbook.worksheets; folhas.load('items/name'); await ctx.sync();
      var usados = folhas.items.map(function (ws) { var u = ws.getUsedRangeOrNullObject(true); u.load('values,formulas,rowIndex,columnIndex'); return { ws: ws, u: u }; });
      await ctx.sync();
      var cells = [];
      usados.forEach(function (x) {
        if (x.u.isNullObject) return;
        x.u.values.forEach(function (linha, i) {
          linha.forEach(function (v, j) {
            var f = x.u.formulas[i][j];
            if (typeof f === 'string' && f[0] === '=') return;
            var t = typeof v === 'string' ? v : (typeof v === 'number' && Number.isInteger(v) && Math.abs(v) >= 1e8 ? String(v) : null);
            if (t && t.trim()) cells.push({ folha: x.ws.name, r: x.u.rowIndex + i, c: x.u.columnIndex + j, t: t });
          });
        });
      });
      return cells;
    });
  }
  async function analisar() {
    if (!C.licenciado()) return;
    var btn = $('btnAnalyze'); C.busy(btn, true, 'A analisar…');
    C.setStatus('analyzeStatus', 'A ler o livro…', 'busy');
    try {
      st.cells = await lerCelulas();
      await C.tick();
      var textos = st.cells.map(function (x) { return x.t; });
      var existentes = E.existingLabels(textos);
      st.applied.forEach(function (en) { var mm = en.label.match(/^\[([A-Z_]+?)_(\d+)\]$/); if (mm) existentes[mm[1]] = Math.max(existentes[mm[1]] || 0, +mm[2]); });
      st.entities = E.analyze(textos, { types: C.tiposActivos(), existingLabels: existentes }).entities;
      C.mostrarEntidades(st.entities, localizar);
      $('review').hidden = false; $('applyCard').hidden = false;
      C.setStatus('analyzeStatus', 'Análise concluída: ' + C.plural(st.cells.length, 'célula com texto lida.', 'células com texto lidas.'), 'ok');
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
    var vars = Array.from(en.variants.keys());
    var x = st.cells.find(function (cl) { return vars.some(function (v) { return cl.t.indexOf(v) >= 0; }); });
    if (!x) return;
    await Excel.run(async function (ctx) { var ws = ctx.workbook.worksheets.getItem(x.folha); ws.activate(); ws.getCell(x.r, x.c).select(); await ctx.sync(); });
  }
  async function escreverCelulas(muda) {
    await Excel.run(async function (ctx) {
      var n = 0;
      for (var i = 0; i < muda.length; i++) {
        var x = muda[i];
        ctx.workbook.worksheets.getItem(x.folha).getCell(x.r, x.c).values = [[x.novo]];
        if (++n % 400 === 0) await ctx.sync();
      }
      await ctx.sync();
    });
  }
  async function anonimizar() {
    if (!C.licenciado()) return;
    var sel = st.entities.filter(function (e) { return e.selected; });
    if (!sel.length) { C.setStatus('applyStatus', 'Não há elementos seleccionados.', 'err'); return; }
    var modo = (document.querySelector('input[name=mode]:checked') || {}).value || 'pseudo';
    var btn = $('btnApply'); C.busy(btn, true, 'A aplicar…');
    try {
      var plano = E.buildPlan(sel, { style: modo === 'x' ? 'x' : 'label' });
      var muda = [];
      st.cells.forEach(function (x) { var novo = E.applyToText(x.t, plano); if (novo !== x.t) muda.push({ folha: x.folha, r: x.r, c: x.c, novo: novo }); });
      await escreverCelulas(muda);
      st.applied = st.applied.concat(sel);
      $('resultSummary').innerHTML = '<p>' + C.plural(muda.length, 'célula alterada.', 'células alteradas.') + '</p>';
      $('tableBlock').hidden = modo !== 'pseudo';
      if (modo === 'pseudo') $('tableText').value = E.buildTable(st.applied, { document: st.livro || 'livro do Excel' });
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
      var cells = await lerCelulas(), muda = [];
      cells.forEach(function (x) { var novo = E.revertText(x.t, mapa); if (novo !== x.t) muda.push({ folha: x.folha, r: x.r, c: x.c, novo: novo }); });
      await escreverCelulas(muda);
      C.setStatus('revertStatus', 'Dados repostos em ' + C.plural(muda.length, 'célula.', 'células.'), 'ok');
    } catch (e) { C.setStatus('revertStatus', C.errMsg(e), 'err'); }
    finally { C.busy(btn, false); }
  }

  window.__cofreExcel = { st: st, deslocarFormula: deslocarFormula, executar: executar };
  boot();
})();
