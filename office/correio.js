/*
 * Cofre para Outlook e Cofre para Gmail: painel comum.
 * Assistente (análise de mensagens, rascunhos de resposta, melhoria do texto) com pseudonimização
 * antes do envio à IA da conta Cloudflare do utilizador, e protecção de dados (verificação, anonimização,
 * pseudonimização e reposição), feita no computador do utilizador.
 * O acesso ao programa de correio é feito por um adaptador: window.CorreioHost.
 * Ana Marta Soares, Especialista em IA aplicada ao Direito.
 */
'use strict';
(function () {
  var C = window.Cofre, E = window.AnonEngine, $ = C.$;
  var H = window.CorreioHost;
  var st = { modo: null, msg: null, ultimo: null, textos: [], entidades: [], aplicadas: [], copia: false };

  /* ============================================================ arranque */
  function boot() {
    if (!window.__cofreSuportado || !E || !C || !H) { $('unsupported').hidden = false; return; }
    C.iniciar({ produto: H.produto, nome: 'Cofre para ' + H.nome, revogadas: 'https://ana-marta-soares-direito-ia.github.io/cofre-anonimizacao-word/revogadas.json' });
    bind();
    H.iniciar().then(function (ok) {
      if (!ok) { $('outside').hidden = false; return; }
      actualizarContexto();
      H.onMudanca(function () { st.ultimo = null; $('resultadoCard').hidden = true; limparProteccao(); actualizarContexto(); });
    }).catch(function () { $('outside').hidden = false; });
  }
  function bind() {
    $('btnExecutar').addEventListener('click', executar);
    $('instrucao').addEventListener('keydown', function (e) { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) executar(); });
    $('btnActualizar').addEventListener('click', actualizarContexto);
    $('btnResponder').addEventListener('click', function () { usarRascunho('responder'); });
    $('btnResponderTodos').addEventListener('click', function () { usarRascunho('todos'); });
    $('btnInserir').addEventListener('click', function () { usarRascunho('inserir'); });
    $('btnCopiar').addEventListener('click', function () { copiar($('rascunhoCorpo').value, 'rascunhoStatus', 'Texto copiado.'); });
    $('btnVerificar').addEventListener('click', verificar);
    $('btnAddTerm').addEventListener('click', acrescentarTermo);
    $('btnApply').addEventListener('click', substituir);
    $('btnCopiaCopiar').addEventListener('click', function () { copiar($('copiaTexto').value, 'copiaStatus', 'Cópia anonimizada copiada.'); });
    $('btnTableCopy').addEventListener('click', function () { copiar($('tableText').value, 'tableStatus', 'Tabela copiada.'); });
    $('btnTableDocx').addEventListener('click', function () { C.tabelaDocx('tabela de correspondencias.docx', $('tableText').value); });
    $('tableFile').addEventListener('change', function () { C.lerFicheiroTabela($('tableFile').files[0]).then(function (t) { $('revertText').value = t; }); });
    $('btnRevert').addEventListener('click', repor);
  }

  /* ============================================================ contexto */
  function actualizarContexto() {
    return H.modo().then(function (m) {
      st.modo = m;
      document.body.dataset.modo = m || 'nenhum';
      var leitura = m === 'leitura', comp = m === 'composicao';
      $('semMensagem').hidden = !!m;
      $('contexto').hidden = !m;
      $('sugLeitura').hidden = !leitura; $('sugComposicao').hidden = !comp;
      $('btnResponder').hidden = !leitura; $('btnResponderTodos').hidden = !leitura; $('btnInserir').hidden = !comp; $('notaInserir').hidden = !comp;
      $('modoTexto').textContent = leitura ? 'Mensagem aberta' : comp ? 'Rascunho em escrita' : '';
      $('instrucao').placeholder = comp
        ? 'Por exemplo: melhora o meu texto, num registo jurídico e cordial.'
        : 'Por exemplo: resume a mensagem e prepara uma resposta a confirmar a reunião.';
      $('applyTitulo').textContent = comp ? 'Substituição no rascunho' : 'Cópia anonimizada';
      $('applyNota').textContent = comp
        ? 'As substituições são feitas no assunto e no corpo do rascunho. Os anexos não são alterados.'
        : 'A mensagem recebida não é alterada: é gerada uma cópia do assunto e do corpo com os dados substituídos, para copiar ou guardar.';
      $('btnApply').textContent = comp ? 'Aplicar ao rascunho' : 'Gerar a cópia';
      $('cardRepor').hidden = !comp;
      if (!m) return;
      return H.lerMensagem().then(function (msg) {
        st.msg = msg;
        $('ctxAssunto').textContent = msg.assunto || '(sem assunto)';
        $('ctxDe').textContent = leitura ? (msg.de || '') : (msg.para && msg.para.length ? 'Para: ' + msg.para.join(', ') : 'Sem destinatários');
        $('ctxAnexos').textContent = msg.anexos && msg.anexos.length ? C.plural(msg.anexos.length, 'anexo', 'anexos') + ': ' + msg.anexos.join(', ') : '';
      });
    }).catch(function (e) { C.setStatus('assistStatus', C.errMsg(e), 'err'); });
  }

  /* ============================================================ assistente */
  function sistema() {
    return 'És o assistente do Cofre para ' + H.nome + ', ao serviço de juristas, advogados, solicitadores e serviços públicos em Portugal. ' +
      'Recebes uma instrução do utilizador e uma mensagem de correio electrónico que ele recebeu, ou o rascunho que está a escrever.\n' +
      'Respondes SEMPRE e APENAS com um objecto JSON, sem texto antes ou depois, com esta forma:\n' +
      '{"resposta":"texto para o utilizador: a análise, o resumo, a resposta à pergunta ou uma frase a apresentar o rascunho",' +
      '"pontos":{"pedidos":["o que é pedido ao utilizador"],"prazos":[{"data":"data ou prazo, tal como resulta da mensagem","descricao":"..."}],"atencao":["pontos jurídicos ou de risco a ponderar"],"prioridade":"alta|media|baixa"},' +
      '"rascunho":{"assunto":"...","corpo":"..."}}\n' +
      'Regras:\n' +
      '1. Preenche "pontos" apenas quando a instrução pede análise, resumo, prazos ou pedidos; caso contrário, "pontos":null.\n' +
      '2. Preenche "rascunho" apenas quando a instrução pede uma resposta, um texto novo ou a melhoria do rascunho; caso contrário, "rascunho":null. O corpo é texto simples, com os parágrafos separados por \\n, com saudação e fórmula de despedida adequadas, sem assinatura com nome (termina na fórmula de despedida). Ao melhorar um rascunho, devolve apenas o texto escrito pelo utilizador, melhorado, sem a mensagem citada.\n' +
      '3. Escreve em português europeu, com a grafia anterior ao Acordo Ortográfico de 1990, em registo profissional e jurídico, sem travessões. Se a mensagem recebida estiver noutra língua, o rascunho é escrito nessa língua, salvo instrução em contrário; a "resposta" é sempre em português.\n' +
      '4. As etiquetas como [PESSOA_1], [EMAIL_1] ou [NIF_1] substituem dados pessoais: mantém-nas exactamente como estão, sem as explicar nem alterar.\n' +
      '5. Não inventes factos, datas, valores, documentos nem compromissos que não constem da mensagem ou da instrução. Quando faltar informação para o rascunho, escreve [a completar].\n' +
      '6. Nos prazos, indica a data ou o prazo tal como resulta da mensagem e não calcules datas que a mensagem não permita determinar.\n' +
      '7. Não prestes aconselhamento definitivo: nos pontos de atenção, assinala o que deve ser verificado pelo utilizador.\n' +
      '8. No rascunho de resposta, a saudação dirige-se ao remetente da mensagem recebida; não repitas o tratamento que o remetente usou para o utilizador.';
  }
  function compor(msg, instr) {
    var tom = $('tom').value;
    var l = ['INSTRUÇÃO: ' + instr];
    if (tom) l.push('TOM PRETENDIDO PARA O RASCUNHO: ' + tom);
    if (st.modo === 'composicao') {
      l.push('', 'RASCUNHO EM ESCRITA', 'Assunto: ' + (msg.assunto || ''), 'Para: ' + (msg.para || []).join(', '));
      if (msg.seleccao) l.push('TEXTO SELECCIONADO PELO UTILIZADOR (a instrução refere-se a este texto):', msg.seleccao, '');
      l.push('CORPO DO RASCUNHO (pode incluir, no fim, a mensagem citada a que se responde):', cortar(msg.corpo, 40000));
    } else {
      l.push('', 'MENSAGEM RECEBIDA', 'Assunto: ' + (msg.assunto || ''), 'De: ' + (msg.de || ''), 'Para: ' + (msg.para || []).join(', '));
      if (msg.cc && msg.cc.length) l.push('Cc: ' + msg.cc.join(', '));
      if (msg.data) l.push('Data: ' + msg.data);
      if (msg.anexos && msg.anexos.length) l.push('Anexos: ' + msg.anexos.join(', '));
      l.push('CORPO:', cortar(msg.corpo, 60000));
    }
    return l.join('\n');
  }
  function cortar(t, n) { t = String(t || ''); return t.length > n ? t.slice(0, n) + '\n[…texto truncado…]' : t; }

  function executar() {
    if (!C.licenciado()) return;
    var instr = $('instrucao').value.trim();
    if (!instr) { C.setStatus('assistStatus', 'Escreva o que pretende ou escolha uma das sugestões.', 'err'); return; }
    if (!st.modo) { C.setStatus('assistStatus', 'Abra uma mensagem ou um rascunho e tente de novo.', 'err'); return; }
    var btn = $('btnExecutar'); C.busy(btn, true, 'A preparar…');
    C.setStatus('assistStatus', 'A ler ' + (st.modo === 'composicao' ? 'o rascunho' : 'a mensagem') + '…', 'busy');
    $('resultadoCard').hidden = true;
    H.lerMensagem().then(function (msg) {
      st.msg = msg;
      var texto = compor(msg, instr);
      var m = C.mascara([texto], $('optMascara').checked);
      C.setStatus('assistStatus', (m.n ? C.plural(m.n, 'dado pessoal substituído', 'dados pessoais substituídos') + ' por etiquetas. ' : '') + 'A aguardar o serviço de IA…', 'busy');
      return C.pedirIA({ sistema: sistema(), mensagem: m.tapar(texto), maxTokens: 5000, esforco: 'medio' }).then(function (t) {
        var j = C.extrairJSON(t);
        return C.destaparFundo(j, m);
      });
    }).then(function (j) {
      st.ultimo = j;
      mostrarResultado(j);
      C.setStatus('assistStatus', '');
    }).catch(function (e) { C.setStatus('assistStatus', e.message || C.errMsg(e), 'err'); })
      .then(function () { C.busy(btn, false); });
  }
  function lista(titulo, itens) {
    if (!itens || !itens.length) return '';
    return '<h3>' + C.esc(titulo) + '</h3><ul class="plano">' + itens.map(function (x) { return '<li>' + C.esc(x) + '</li>'; }).join('') + '</ul>';
  }
  function mostrarResultado(j) {
    $('resposta').textContent = j.resposta || '';
    var p = j.pontos, h = '';
    if (p && typeof p === 'object') {
      var pr = String(p.prioridade || '').toLowerCase();
      if (pr) h += '<p class="prio prio-' + (pr.indexOf('alta') >= 0 ? 'alta' : pr.indexOf('baixa') >= 0 ? 'baixa' : 'media') + '">Prioridade ' + C.esc(pr.replace('media', 'média')) + '</p>';
      h += lista('Pedidos', p.pedidos);
      h += lista('Prazos', (p.prazos || []).map(function (x) { return typeof x === 'string' ? x : (x.data ? x.data + ': ' : '') + (x.descricao || ''); }));
      h += lista('Pontos de atenção', p.atencao);
    }
    $('pontos').innerHTML = h; $('pontos').hidden = !h;
    var r = j.rascunho && (j.rascunho.corpo || j.rascunho.assunto) ? j.rascunho : null;
    $('rascunhoBloco').hidden = !r;
    if (r) { $('rascunhoAssunto').value = r.assunto || ''; $('rascunhoCorpo').value = r.corpo || ''; }
    C.setStatus('rascunhoStatus', '');
    $('resultadoCard').hidden = false;
    $('resultadoCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function usarRascunho(como) {
    var corpo = $('rascunhoCorpo').value.trim();
    if (!corpo) { C.setStatus('rascunhoStatus', 'O rascunho está vazio.', 'err'); return; }
    var p = como === 'inserir' ? H.inserir(corpo) : H.responder(corpo, como === 'todos');
    C.setStatus('rascunhoStatus', 'A transferir o texto…', 'busy');
    p.then(function (msg) {
      C.setStatus('rascunhoStatus', msg || (como === 'inserir' ? 'Texto inserido no rascunho. Reveja-o antes de enviar.' : 'Foi aberta a resposta com o texto. Reveja-a antes de enviar.'), 'ok');
    }).catch(function (e) {
      copiar(corpo, null);
      C.setStatus('rascunhoStatus', (e && e.message ? e.message + ' ' : '') + 'O texto foi copiado: cole-o na resposta com Cmd+V ou Ctrl+V.', 'err');
    });
  }
  function copiar(texto, statusId, msgOk) {
    var feito = function () { if (statusId) C.setStatus(statusId, msgOk, 'ok'); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(texto).then(feito).catch(function () { copiarAntigo(texto) ? feito() : statusId && C.setStatus(statusId, 'Seleccione o texto e copie-o manualmente.', 'err'); });
    }
    if (copiarAntigo(texto)) feito(); else if (statusId) C.setStatus(statusId, 'Seleccione o texto e copie-o manualmente.', 'err');
  }
  function copiarAntigo(texto) {
    var ta = document.createElement('textarea'); ta.value = texto; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    var ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove(); return ok;
  }

  /* ============================================================ protecção de dados */
  // Indícios de categorias especiais de dados (artigo 9.º do RGPD) e de dados penais (artigo 10.º)
  var ESPECIAIS = [
    ['saúde', /\b(diagn[óo]stic\w*|doen[çc]a\w*|baixa m[ée]dica|atestado m[ée]dico|relat[óo]rio m[ée]dico|incapacidade|defici[êe]ncia|gravidez|gr[áa]vida|tratamento m[ée]dico|medica[çc][ãa]o|internamento|oncol[óo]gic\w*|psiqui[áa]tric\w*|psicol[óo]gic\w*|depress[ãa]o|cl[íi]nic[oa]s?)\b/i],
    ['origem racial ou étnica', /\b(origem (racial|[ée]tnica)|etnia|ra[çc]a)\b/i],
    ['convicções religiosas ou filosóficas', /\b(religi[ãa]o|religios[oa]s?|confiss[ãa]o religiosa|cren[çc]as?)\b/i],
    ['filiação sindical', /\b(sindicat\w*|filia[çc][ãa]o sindical|sindicalizad[oa])\b/i],
    ['opiniões políticas', /\b(filia[çc][ãa]o partid[áa]ria|partido pol[íi]tico|opini[õo]es pol[íi]ticas)\b/i],
    ['vida sexual ou orientação sexual', /\b(orienta[çc][ãa]o sexual|vida sexual|identidade de g[ée]nero)\b/i],
    ['dados genéticos ou biométricos', /\b(gen[ée]tic\w*|biom[ée]tric\w*|impress[ãa]o digital|ADN)\b/i],
    ['condenações penais e infracções (artigo 10.º)', /\b(registo criminal|cadastro|condena[çc][ãa]o|condenad[oa]|arguid[oa]|crime|pena de pris[ãa]o|processo[- ]crime|inqu[ée]rito criminal)\b/i]
  ];
  function textosMsg(msg) {
    var l = [msg.assunto || ''];
    String(msg.corpo || '').split(/\n/).forEach(function (x) { l.push(x); });
    return l.filter(function (x) { return x.trim(); });
  }
  function limparProteccao() {
    st.textos = []; st.entidades = [];
    ['review', 'applyCard', 'resultCard', 'alertas'].forEach(function (id) { $(id).hidden = true; });
    C.setStatus('verificarStatus', ''); C.setStatus('applyStatus', '');
  }
  function verificar() {
    if (!C.licenciado()) return;
    if (!st.modo) { C.setStatus('verificarStatus', 'Abra uma mensagem ou um rascunho e tente de novo.', 'err'); return; }
    var btn = $('btnVerificar'); C.busy(btn, true, 'A verificar…');
    C.setStatus('verificarStatus', 'A analisar neste computador…', 'busy');
    var leitura = st.modo === 'composicao' ? H.lerNos() : H.lerMensagem().then(function (msg) { st.msg = msg; return textosMsg(msg); });
    leitura.then(function (textos) {
      return C.tick().then(function () { return textos; });
    }).then(function (textos) {
      st.textos = textos;
      var existentes = E.existingLabels(textos);
      st.aplicadas.forEach(function (en) { var mm = en.label.match(/^\[([A-Z_]+?)_(\d+)\]$/); if (mm) existentes[mm[1]] = Math.max(existentes[mm[1]] || 0, +mm[2]); });
      st.entidades = E.analyze(textos, { types: C.tiposActivos(), existingLabels: existentes }).entities;
      C.mostrarEntidades(st.entidades, null);
      return H.lerMensagem().then(function (msg) { mostrarAlertas(textos.join('\n'), msg); });
    }).then(function () {
      $('review').hidden = false; $('applyCard').hidden = false; $('resultCard').hidden = true;
      C.setStatus('verificarStatus', 'Verificação concluída. Nada foi enviado para fora deste computador.', 'ok');
    }).catch(function (e) { C.setStatus('verificarStatus', C.errMsg(e), 'err'); })
      .then(function () { C.busy(btn, false); });
  }
  function dominio(endereco) { var m = String(endereco || '').match(/@([^>\s,;]+)/); return m ? m[1].toLowerCase() : ''; }
  function mostrarAlertas(texto, msg) {
    var a = [];
    var cats = ESPECIAIS.filter(function (c) { return c[1].test(texto); }).map(function (c) { return c[0]; });
    if (cats.length) a.push(['danger', 'A mensagem parece conter categorias especiais de dados pessoais ou dados relativos a infracções: ' + cats.join('; ') + '. O tratamento destes dados está sujeito aos artigos 9.º e 10.º do RGPD; confirme a necessidade da sua inclusão e a segurança do meio de envio.']);
    if (st.modo === 'composicao') {
      var vis = (msg.para || []).concat(msg.cc || []);
      var doms = {}; vis.forEach(function (x) { var d = dominio(x); if (d) doms[d] = 1; });
      if (vis.length >= 5 && Object.keys(doms).length > 1) a.push(['warn', 'Há ' + vis.length + ' destinatários visíveis, de ' + Object.keys(doms).length + ' entidades diferentes. Cada destinatário verá os endereços dos restantes: considere colocá-los em Bcc.']);
      if (msg.anexos && msg.anexos.length) a.push(['info', 'O rascunho tem ' + C.plural(msg.anexos.length, 'anexo', 'anexos') + ' (' + msg.anexos.join(', ') + '). Os anexos não são verificados: confirme que não contêm dados pessoais desnecessários.']);
    }
    var box = $('alertas'); box.innerHTML = a.map(function (x) { return '<div class="notice small ' + x[0] + '">' + C.esc(x[1]) + '</div>'; }).join('');
    box.hidden = !a.length;
  }
  function acrescentarTermo() {
    var term = $('manualTerm').value.trim(); if (!term) return;
    var en = E.addManual(st.entidades.concat(st.aplicadas), term, $('manualType').value);
    if (en && st.entidades.indexOf(en) < 0 && st.aplicadas.indexOf(en) < 0) st.entidades.push(en);
    $('manualTerm').value = '';
    C.mostrarEntidades(st.entidades, null);
  }
  function substituir() {
    if (!C.licenciado()) return;
    var sel = st.entidades.filter(function (e) { return e.selected; });
    if (!sel.length) { C.setStatus('applyStatus', 'Não há elementos seleccionados.', 'err'); return; }
    var modo = (document.querySelector('input[name=mode]:checked') || {}).value || 'pseudo';
    var plano = E.buildPlan(sel, { style: modo === 'x' ? 'x' : 'label' });
    var btn = $('btnApply'); C.busy(btn, true, 'A aplicar…');
    var feito;
    if (st.modo === 'composicao') {
      feito = H.lerNos().then(function (textos) {
        var novos = textos.map(function (t) { return E.applyToText(t, plano); });
        var n = novos.filter(function (t, i) { return t !== textos[i]; }).length;
        return H.escreverNos(novos).then(function () { return n; });
      }).then(function (n) {
        $('resultSummary').innerHTML = '<p>' + C.plural(n, 'parágrafo alterado no rascunho.', 'parágrafos alterados no rascunho.') + '</p>';
        $('copiaBloco').hidden = true;
      });
    } else {
      feito = H.lerMensagem().then(function (msg) {
        var assunto = E.applyToText(msg.assunto || '', plano);
        var corpo = String(msg.corpo || '').split('\n').map(function (l) { return E.applyToText(l, plano); }).join('\n');
        $('copiaTexto').value = 'Assunto: ' + assunto + '\n\n' + corpo;
        $('copiaBloco').hidden = false;
        $('resultSummary').innerHTML = '<p>Cópia gerada. A mensagem original não foi alterada.</p>';
      });
    }
    feito.then(function () {
      st.aplicadas = st.aplicadas.concat(sel);
      $('tableBlock').hidden = modo !== 'pseudo';
      if (modo === 'pseudo') $('tableText').value = E.buildTable(st.aplicadas, { document: (st.msg && st.msg.assunto) || 'mensagem de correio electrónico' });
      $('resultCard').hidden = false;
      C.setStatus('applyStatus', 'Substituição concluída.', 'ok');
      st.entidades = st.entidades.filter(function (e) { return !e.selected; });
      C.mostrarEntidades(st.entidades, null);
      $('resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }).catch(function (e) { C.setStatus('applyStatus', C.errMsg(e), 'err'); })
      .then(function () { C.busy(btn, false); });
  }
  function repor() {
    if (!C.licenciado()) return;
    var mapa = E.parseTable($('revertText').value);
    if (!mapa.length) { C.setStatus('revertStatus', 'A tabela não contém correspondências válidas.', 'err'); return; }
    var btn = $('btnRevert'); C.busy(btn, true, 'A repor…');
    H.lerNos().then(function (textos) {
      var novos = textos.map(function (t) { return E.revertText(t, mapa); });
      var n = novos.filter(function (t, i) { return t !== textos[i]; }).length;
      return H.escreverNos(novos).then(function () { C.setStatus('revertStatus', 'Dados repostos em ' + C.plural(n, 'parágrafo.', 'parágrafos.'), 'ok'); });
    }).catch(function (e) { C.setStatus('revertStatus', C.errMsg(e), 'err'); })
      .then(function () { C.busy(btn, false); });
  }

  window.__cofreCorreio = { st: st, executar: executar };
  boot();
})();
