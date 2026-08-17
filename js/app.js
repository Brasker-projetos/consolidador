/* app.js — orquestração da UI e do pipeline (§5,§8,§11). */
(function () {
  'use strict';
  const U = window.U, P = window.P, C = window.C, X = window.X;
  const files = { balA: null, dreA: null, balB: null, dreB: null };
  let empA = null, empB = null, detected = null, consState = null;

  // ---------- Upload ----------
  document.querySelectorAll('input[type=file]').forEach(inp => {
    inp.addEventListener('change', e => {
      const k = inp.dataset.k;
      const f = e.target.files[0];
      if (!f) return;
      files[k] = f;
      const drop = inp.closest('.drop');
      drop.classList.add('ok');
      drop.querySelector('.fn').textContent = f.name;
      checkReady();
    });
  });

  function checkReady() {
    const ready = files.balA && files.dreA && files.balB && files.dreB;
    document.getElementById('btnAnalisar').disabled = !ready;
    document.getElementById('uploadHint').textContent = ready ? 'Pronto para analisar.' :
      'Faltam arquivos (' + Object.values(files).filter(Boolean).length + '/4).';
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
  document.getElementById('btnAnalisar').addEventListener('click', async () => {
    const msg = document.getElementById('analiseMsg');
    msg.innerHTML = '<div class="msg"><span class="spin"></span>Lendo e parseando os 4 arquivos…</div>';
    try {
      const [gBA, gDA, gBB, gDB] = await Promise.all([
        readGrid(files.balA), readGrid(files.dreA),
        readGrid(files.balB), readGrid(files.dreB)
      ]);
      empA = { id: 'A', balanco: P.parseBalance(gBA, files.balA.name), dre: P.parseDRE(gDA, files.dreA.name) };
      empB = { id: 'B', balanco: P.parseBalance(gBB, files.balB.name), dre: P.parseDRE(gDB, files.dreB.name) };
      msg.innerHTML = '';
      renderAnalise();
      detected = C.detectIntercompany(empA, empB);
      renderInter();
      document.getElementById('cardAnalise').classList.remove('hidden');
      document.getElementById('cardInter').classList.remove('hidden');
      document.getElementById('cardGerar').classList.remove('hidden');
    } catch (err) {
      msg.innerHTML = '<div class="msg err"><b>Erro:</b> ' + err.message + '</div>';
      console.error(err);
    }
  });

  function fechaTag(b) {
    if (b === null) return '<span class="tag warn">sem total</span>';
    return b ? '<span class="tag ok">✓ fecha</span>' : '<span class="tag err">✗ não fecha</span>';
  }

  function renderAnalise() {
    const tb = document.querySelector('#tblEmpresas tbody');
    tb.innerHTML = '';
    [['Empresa A', empA], ['Empresa B', empB]].forEach(([rot, e]) => {
      const h = e.balanco.header;
      tb.insertAdjacentHTML('beforeend',
        `<tr><td><b>${rot}</b><br><small>${escape(h.nome || '—')}</small></td>
         <td>${h.cnpjFmt || '—'}</td><td>${h.data || '—'}</td>
         <td class="num">${U.fmt(Math.abs(e.balanco.totalAtivo))}</td>
         <td class="num">${U.fmt(Math.abs(e.balanco.totalPassivo))}</td>
         <td>${fechaTag(e.balanco.fecha)}</td></tr>`);
    });
    // checagem DRE
    const chk = [];
    [['A', empA], ['B', empB]].forEach(([rot, e]) => {
      const probs = e.dre.sanity || [];
      chk.push(`Empresa ${rot}: ` + (probs.length ? '⚠ ' + probs.join('; ') : '✓ DRE lida e consistente') +
        ` — Result. Líquido = ${U.fmt(e.dre.resultLiquido)}`);
    });
    document.getElementById('dreCheck').innerHTML = chk.join('<br>');
  }

  // ---------- Intercompany ----------
  function renderInter() {
    const tb = document.querySelector('#tblInter tbody');
    tb.innerHTML = '';
    const pairs = detected.pairs;
    if (!pairs.length) {
      document.getElementById('interVazio').classList.remove('hidden');
      return;
    }
    document.getElementById('interVazio').classList.add('hidden');
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
          <td><small>${escape(p.a.desc)}</small></td>
          <td class="num">${U.fmt(p.valA)}</td>
          <td><small>${escape(p.b.desc)}</small></td>
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
  document.getElementById('btnGerar').addEventListener('click', async () => {
    const out = document.getElementById('resultado');
    out.innerHTML = '<div class="msg"><span class="spin"></span>Consolidando e validando…</div>';
    try {
      const cons = C.consolidateBalance(empA, empB, detected.pairs);
      const dcons = C.consolidateDRE(empA, empB);
      const positions = C.financialPositions(empA, empB, detected.pairs);
      consState = { cons, dcons, positions };

      // validações (§8)
      const blocks = [];
      if (empA.balanco.fecha === false) blocks.push('Balanço da Empresa A não fecha na origem.');
      if (empB.balanco.fecha === false) blocks.push('Balanço da Empresa B não fecha na origem.');
      if (!cons.fecha) blocks.push('Balanço consolidado não fecha: diferença R$ ' + U.fmt(cons.diff) + '.');
      (dcons.errors || []).forEach(e => blocks.push(e));

      renderResultado(cons, dcons, blocks, cons.residual, positions);
    } catch (err) {
      out.innerHTML = '<div class="msg err"><b>Erro:</b> ' + err.message + '</div>';
      console.error(err);
    }
  });

  function renderResultado(cons, dcons, blocks, residual, positions) {
    const out = document.getElementById('resultado');
    let html = `<div class="resume">
      <div><span>Total Ativo consolidado</span><b>${U.fmt(cons.totalAtivo)}</b></div>
      <div><span>Total Passivo + PL</span><b>${U.fmt(cons.totalPassivo)}</b></div>
      <div><span>Resultado Líquido</span><b>${U.fmt(dcons.resultLiquido)}</b></div>
    </div>`;
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
        blocks.map(b => '<li>' + b + '</li>').join('') + '</ul>' +
        'A versão oficial não deve ser usada até a correção na origem ou o ajuste do tratamento da divergência.</div>';
      html += '<button class="sec" id="btnForce">Baixar versão provisória (com pendências)</button>';
    } else {
      html += '<div class="msg ok">✓ Todas as validações passaram. Pronto para baixar.</div>';
      html += '<button id="btnBaixar">Baixar .xlsx consolidado</button>';
    }
    out.innerHTML = html;
    const bd = document.getElementById('btnBaixar'); if (bd) bd.addEventListener('click', baixar);
    const bf = document.getElementById('btnForce'); if (bf) bf.addEventListener('click', baixar);
  }

  async function baixar(ev) {
    const btn = ev.currentTarget;
    const provisoria = btn.id === 'btnForce';
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span>Gerando…';
    try {
      const fileNames = [files.balA.name, files.dreA.name, files.balB.name, files.dreB.name];
      const blob = await X.generate(empA, empB, consState.cons, consState.dcons, detected, fileNames);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = (provisoria ? 'PROVISORIO_' : 'Consolidado_') +
        (empA.balanco.header.data || '').replace(/\//g, '-') + '.xlsx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      btn.disabled = false; btn.textContent = provisoria ?
        '✓ Versão provisória baixada — gerar de novo' : '✓ Baixado — gerar de novo';
    } catch (err) {
      btn.disabled = false; btn.textContent = 'Erro — tentar de novo';
      alert('Erro ao gerar: ' + err.message); console.error(err);
    }
  }

  function escape(s) { return String(s == null ? '' : s).replace(/[&<>]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[m])); }
})();
