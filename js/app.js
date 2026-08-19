/* app.js — orquestração da UI e do pipeline (§5,§8,§11).
   A quantidade de empresas é escolhida na tela (2 a 10); cada empresa tem um
   par de arquivos (Balanço e DRE) e vira uma coluna no relatório final. */
(function () {
  'use strict';
  const U = window.U, P = window.P, C = window.C, X = window.X;
  const MAX = (C && C.MAX_EMPRESAS) || 10;
  const MIN = (C && C.MIN_EMPRESAS) || 2;
  const ID = 'ABCDEFGHIJ';

  // Um slot por empresa possível. Reduzir a quantidade não apaga o que já foi
  // carregado: os arquivos voltam se a pessoa aumentar de novo.
  const slots = [];
  for (let i = 0; i < MAX; i++) slots.push({ bal: null, dre: null });
  let qtd = MIN;
  let emps = null, detected = null, consState = null;

  const $ = id => document.getElementById(id);
  function escape(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m])); }

  // ---------- Quantidade de empresas ----------
  const selQtd = $('qtdEmpresas');
  for (let n = MIN; n <= MAX; n++) {
    selQtd.insertAdjacentHTML('beforeend', '<option value="' + n + '">' + n + ' empresas</option>');
  }
  selQtd.value = String(qtd);
  selQtd.addEventListener('change', () => {
    const n = parseInt(selQtd.value, 10);
    qtd = Math.min(MAX, Math.max(MIN, isFinite(n) ? n : MIN));
    selQtd.value = String(qtd);
    resetAnalise();
    renderSlots();
  });

  // ---------- Upload ----------
  function dropHtml(i, tipo, rotulo, file) {
    return '<label class="drop' + (file ? ' ok' : '') + '">' +
      '<b>' + rotulo + ' — Empresa ' + ID[i] + '</b>' +
      '<small>.xls ou .xlsx</small>' +
      '<span class="fn">' + (file ? escape(file.name) : '') + '</span>' +
      '<input type="file" accept=".xls,.xlsx" data-i="' + i + '" data-t="' + tipo + '"></label>';
  }

  function renderSlots() {
    let html = '';
    for (let i = 0; i < qtd; i++) {
      html += '<div class="empresa"><h3>Empresa ' + ID[i] + '</h3><div class="grid">' +
        dropHtml(i, 'bal', 'Balanço', slots[i].bal) +
        dropHtml(i, 'dre', 'DRE', slots[i].dre) +
        '</div></div>';
    }
    const box = $('empresasSlots');
    box.innerHTML = html;
    box.querySelectorAll('input[type=file]').forEach(inp => inp.addEventListener('change', onPick));
    checkReady();
  }

  function onPick(e) {
    const inp = e.target, f = inp.files && inp.files[0];
    if (!f) return;
    slots[+inp.dataset.i][inp.dataset.t] = f;
    const drop = inp.closest('.drop');
    drop.classList.add('ok');
    drop.querySelector('.fn').textContent = f.name;
    // Arquivo novo invalida a análise anterior: sem isso dava para gerar a
    // consolidação com o resultado antigo, de outros arquivos.
    resetAnalise();
    checkReady();
  }

  function checkReady() {
    let carregados = 0;
    for (let i = 0; i < qtd; i++) { if (slots[i].bal) carregados++; if (slots[i].dre) carregados++; }
    const total = qtd * 2;
    $('btnAnalisar').disabled = carregados < total;
    $('uploadHint').textContent = carregados >= total ? 'Pronto para analisar.' :
      'Faltam arquivos (' + carregados + '/' + total + ').';
  }

  function resetAnalise() {
    emps = null; detected = null; consState = null;
    ['cardAnalise', 'cardInter', 'cardGerar'].forEach(id => $(id).classList.add('hidden'));
    $('resultado').innerHTML = '';
    $('analiseMsg').innerHTML = '';
  }

  function readGrid(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => {
        try {
          const grid = window.LOAD.gridFromBuffer(new Uint8Array(fr.result));
          if (!grid || !grid.length) throw new Error('arquivo sem dados legíveis');
          resolve(grid);
        } catch (err) { reject(new Error('Falha ao ler ' + file.name + ': ' + err.message)); }
      };
      fr.onerror = () => reject(new Error('Erro de leitura em ' + file.name));
      fr.readAsArrayBuffer(file);
    });
  }

  // ---------- Analisar ----------
  $('btnAnalisar').addEventListener('click', async () => {
    const msg = $('analiseMsg');
    const total = qtd * 2;
    msg.innerHTML = '<div class="msg"><span class="spin"></span>Lendo e parseando os ' + total + ' arquivos…</div>';
    try {
      const pedidos = [];
      for (let i = 0; i < qtd; i++) pedidos.push(readGrid(slots[i].bal), readGrid(slots[i].dre));
      const grids = await Promise.all(pedidos);
      const lidas = [];
      for (let i = 0; i < qtd; i++) {
        lidas.push({
          id: ID[i],
          balanco: P.parseBalance(grids[i * 2], slots[i].bal.name),
          dre: P.parseDRE(grids[i * 2 + 1], slots[i].dre.name)
        });
      }
      emps = lidas;
      msg.innerHTML = '';
      renderAnalise();
      detected = C.detectIntercompany(emps);
      renderInter();
      $('cardAnalise').classList.remove('hidden');
      $('cardInter').classList.remove('hidden');
      $('cardGerar').classList.remove('hidden');
    } catch (err) {
      resetAnalise();
      msg.innerHTML = '<div class="msg err"><b>Erro:</b> ' + escape(err.message) + '</div>';
      console.error(err);
    }
  });

  function fechaTag(b) {
    if (b === null) return '<span class="tag warn">sem total</span>';
    return b ? '<span class="tag ok">✓ fecha</span>' : '<span class="tag err">✗ não fecha</span>';
  }

  // Conferências que só aparecem quando se carrega várias empresas de uma vez:
  // o mesmo arquivo em dois slots, ou balanços de datas diferentes.
  function avisosDeEntrada() {
    const avisos = [];
    const porCnpj = new Map();
    emps.forEach(e => {
      const c = U.digits(e.balanco.header.cnpj);
      if (!c) return;
      if (!porCnpj.has(c)) porCnpj.set(c, []);
      porCnpj.get(c).push('Empresa ' + e.id);
    });
    porCnpj.forEach((lista, c) => {
      if (lista.length > 1) avisos.push({ erro: true,
        txt: 'CNPJ repetido (' + c + ') em ' + lista.join(' e ') + ' — o mesmo balanço parece ter sido carregado mais de uma vez.' });
    });
    const datas = [];
    emps.forEach(e => { const d = e.balanco.header.data; if (d && datas.indexOf(d) < 0) datas.push(d); });
    if (datas.length > 1) avisos.push({ erro: false,
      txt: 'Os balanços têm datas diferentes (' + datas.join(', ') + '). Confira se todos são da mesma data-base antes de usar o consolidado.' });
    return avisos;
  }

  function renderAnalise() {
    const tb = document.querySelector('#tblEmpresas tbody');
    tb.innerHTML = '';
    emps.forEach(e => {
      const h = e.balanco.header;
      tb.insertAdjacentHTML('beforeend',
        `<tr><td><b>Empresa ${e.id}</b><br><small>${escape(h.nome || '—')}</small></td>
         <td>${escape(h.cnpjFmt || '—')}</td><td>${escape(h.data || '—')}</td>
         <td class="num">${U.fmt(Math.abs(e.balanco.totalAtivo))}</td>
         <td class="num">${U.fmt(Math.abs(e.balanco.totalPassivo))}</td>
         <td>${fechaTag(e.balanco.fecha)}</td></tr>`);
    });
    // checagem DRE
    const chk = [];
    emps.forEach(e => {
      const probs = e.dre.sanity || [];
      chk.push(`Empresa ${e.id}: ` + (probs.length ? '⚠ ' + escape(probs.join('; ')) : '✓ DRE lida e consistente') +
        ` — Result. Líquido = ${U.fmt(e.dre.resultLiquido)}`);
    });
    avisosDeEntrada().forEach(a => {
      chk.push((a.erro ? '<b style="color:#C0392B">✗ ' : '⚠ ') + escape(a.txt) + (a.erro ? '</b>' : ''));
    });
    $('dreCheck').innerHTML = chk.join('<br>');
  }

  // ---------- Intercompany ----------
  function renderInter() {
    const tb = document.querySelector('#tblInter tbody');
    tb.innerHTML = '';
    const pairs = detected.pairs;
    if (!pairs.length) {
      $('interVazio').classList.remove('hidden');
      return;
    }
    $('interVazio').classList.add('hidden');
    pairs.forEach((p, i) => {
      const div = p.divergencia > 0.02;
      const tratSel = div ?
        `<select data-trat="${i}">
           <option value="residual">Eliminar recíproco (R$ ${U.fmt(p.confirmavel)}) e manter residual visível</option>
           <option value="cheio">Eliminar valor cheio de cada lado</option>
           <option value="manter">Não eliminar / manter ambos</option>
         </select>` : '<span class="sub">recíproco exato</span>';
      tb.insertAdjacentHTML('beforeend',
        `<tr class="${div ? 'diverg' : ''}">
          <td style="text-align:center"><input type="checkbox" data-elim="${i}" checked></td>
          <td><b>Empresa ${escape(p.a.company)}</b><br><small>${escape(p.a.desc)}</small></td>
          <td class="num">${U.fmt(p.valA)}</td>
          <td><b>Empresa ${escape(p.b.company)}</b><br><small>${escape(p.b.desc)}</small></td>
          <td class="num">${U.fmt(p.valB)}</td>
          <td class="num">${div ? '<b style="color:#BF8F00">' + U.fmt(p.divergencia) + '</b>' : '0,00'}</td>
          <td>${tratSel}</td>
        </tr>`);
    });
    tb.querySelectorAll('[data-elim]').forEach(cb => cb.addEventListener('change', () => {
      detected.pairs[+cb.dataset.elim].decisao = cb.checked ? 'eliminar' : 'manter';
    }));
    tb.querySelectorAll('[data-trat]').forEach(sel => sel.addEventListener('change', () => {
      detected.pairs[+sel.dataset.trat].tratamento = sel.value;
    }));
  }

  // ---------- Gerar ----------
  $('btnGerar').addEventListener('click', async () => {
    const out = $('resultado');
    out.innerHTML = '<div class="msg"><span class="spin"></span>Consolidando e validando…</div>';
    try {
      if (!emps || !detected) throw new Error('Analise os arquivos antes de consolidar.');
      const cons = C.consolidateBalance(emps, detected.pairs);
      const dcons = C.consolidateDRE(emps);
      const positions = C.financialPositions(emps, detected.pairs);
      consState = { cons, dcons, positions };

      // validações (§8)
      const blocks = [];
      emps.forEach(e => {
        if (e.balanco.fecha === false) blocks.push('Balanço da Empresa ' + e.id + ' não fecha na origem.');
        // "sem total" não pode passar: sem a linha de total do Ativo/Passivo a
        // empresa entra ZERADA nas linhas de total do consolidado, enquanto as
        // contas dela continuam aparecendo no corpo do relatório.
        else if (e.balanco.fecha === null) blocks.push('Balanço da Empresa ' + e.id +
          ': não foi possível ler o total do Ativo e do Passivo. Essa empresa entraria zerada ' +
          'nas linhas de total do consolidado — verifique o arquivo antes de usar.');
      });
      if (!cons.fecha) blocks.push('Balanço consolidado não fecha: diferença R$ ' + U.fmt(cons.diff) + '.');
      avisosDeEntrada().forEach(a => { if (a.erro) blocks.push(a.txt); });
      (dcons.errors || []).forEach(e => blocks.push(e));

      renderResultado(cons, dcons, blocks, cons.residual, positions);
    } catch (err) {
      out.innerHTML = '<div class="msg err"><b>Erro:</b> ' + escape(err.message) + '</div>';
      console.error(err);
    }
  });

  function renderResultado(cons, dcons, blocks, residual, positions) {
    const out = $('resultado');
    let html = `<div class="resume">
      <div><span>Total Ativo consolidado</span><b>${U.fmt(cons.totalAtivo)}</b></div>
      <div><span>Total Passivo + PL</span><b>${U.fmt(cons.totalPassivo)}</b></div>
      <div><span>Resultado Líquido</span><b>${U.fmt(dcons.resultLiquido)}</b></div>
    </div>`;
    html += `<p class="sub">${emps.length} empresas consolidadas — uma coluna para cada uma na planilha.</p>`;
    html += `<p class="sub">Diferença Ativo − (Passivo+PL): <b>${U.fmt(cons.diff)}</b> ` + fechaTag(cons.fecha) + '</p>';
    if (residual > 0.02) html += `<p class="sub">⚠ Saldo intercompany residual não conciliado destacado na planilha: <b>R$ ${U.fmt(residual)}</b></p>`;
    html += `<p class="sub"><b>Posições financeiras consolidadas</b> — clientes, fornecedores e bancos são destacados no próprio Balanço Consolidado.</p>` +
      `<div class="resume">` +
      `<div><span>Clientes em aberto</span><b>${U.fmt(positions.clientes.total)}</b></div>` +
      `<div><span>Fornecedores em aberto</span><b>${U.fmt(positions.fornecedores.total)}</b></div>` +
      `<div><span>Bancos</span><b>${U.fmt(positions.bancos.total)}</b></div>` +
      `</div>`;
    (dcons.warnings || []).forEach(w => { html += `<p class="sub">⚠ ${escape(w)}</p>`; });

    if (blocks.length) {
      html += '<div class="msg err"><b>Inconsistências encontradas — revisão necessária:</b><ul>' +
        blocks.map(b => '<li>' + escape(b) + '</li>').join('') + '</ul>' +
        'A versão oficial não deve ser usada até a correção na origem ou o ajuste do tratamento da divergência.</div>';
      html += '<button class="sec" id="btnForce">Baixar versão provisória (com pendências)</button>';
    } else {
      html += '<div class="msg ok">✓ Todas as validações passaram. Pronto para baixar.</div>';
      html += '<button id="btnBaixar">Baixar .xlsx consolidado</button>';
    }
    out.innerHTML = html;
    const bd = $('btnBaixar'); if (bd) bd.addEventListener('click', baixar);
    const bf = $('btnForce'); if (bf) bf.addEventListener('click', baixar);
  }

  async function baixar(ev) {
    const btn = ev.currentTarget;
    const provisoria = btn.id === 'btnForce';
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span>Gerando…';
    try {
      const fileNames = [];
      for (let i = 0; i < emps.length; i++) fileNames.push(slots[i].bal.name, slots[i].dre.name);
      const blob = await X.generate(emps, consState.cons, consState.dcons, detected, fileNames);
      const url = URL.createObjectURL(blob);
      const data = emps.map(e => e.balanco.header.data).filter(Boolean)[0] || '';
      const a = document.createElement('a');
      a.href = url; a.download = (provisoria ? 'PROVISORIO_' : 'Consolidado_') +
        data.replace(/\//g, '-') + '.xlsx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      btn.disabled = false; btn.textContent = provisoria ?
        '✓ Versão provisória baixada — gerar de novo' : '✓ Baixado — gerar de novo';
    } catch (err) {
      btn.disabled = false; btn.textContent = 'Erro — tentar de novo';
      alert('Erro ao gerar: ' + err.message); console.error(err);
    }
  }

  renderSlots();
})();
