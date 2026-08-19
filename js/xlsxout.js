/* xlsxout.js — gera o .xlsx consolidado no layout da referência (ExcelJS).
   O número de colunas acompanha o número de empresas consolidadas:
   Descrição | uma coluna por empresa | Soma | Eliminações | Consolidado. */
(function (global) {
  'use strict';
  const U = global.U;

  const NUMFMT = '#,##0.00;(#,##0.00);"-"';
  const COL = {
    titulo: 'FF1F3864', raiz: 'FFD9E1F2', section: 'FFF2F2F2', total: 'FFBDD7EE',
    grand: 'FF9DC3E6', inter: 'FFFCE4D6', alerta: 'FFFFF2CC', branco: 'FFFFFFFF',
    txtInter: 'FFC55A11', txtAlerta: 'FFBF8F00'
  };
  const FONT = 'Arial';
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = () => ({ style: 'thin', color: { argb: 'FFD9D9D9' } });
  const borderAll = () => ({ top: thin(), bottom: thin(), left: thin(), right: thin() });
  const round2 = n => Math.round((Number(n) || 0) * 100) / 100;
  const firstWord = s => String(s || '').trim().split(/\s+/)[0] || 'Empresa';

  // Letra da coluna no Excel (1 → A, 2 → B…). Com 10 empresas o relatório
  // termina na coluna N, então uma lista fixa de letras não bastaria.
  function colLetter(n) {
    let s = '';
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // Onde fica cada coluna, dado o número de empresas:
  // 1 = Descrição · 2..n+1 = empresas · n+2 = Soma · n+3 = Eliminações · n+4 = Consolidado
  function layoutFor(n) {
    return { n: n, first: 2, soma: n + 2, elim: n + 3, cons: n + 4, last: n + 4 };
  }

  // Rótulo curto de cada empresa para o cabeçalho da coluna: a primeira palavra
  // que só aparece no nome dela. Rótulos repetidos deixariam duas colunas com o
  // mesmo título, então o id da empresa entra como desempate.
  function companyLabels(nomes, ids) {
    const tokens = s => U.norm(s).split(/\s+/).filter(x => x && !/^(LTDA|EIRELI|S\/A|SA)$/.test(x));
    const toks = nomes.map(tokens);
    const labels = nomes.map((nome, i) => {
      const outros = new Set();
      toks.forEach((t, j) => { if (j !== i) t.forEach(x => outros.add(x)); });
      return toks[i].find(x => !outros.has(x)) || firstWord(nome);
    });
    const conta = {};
    labels.forEach(l => { conta[l] = (conta[l] || 0) + 1; });
    return labels.map((l, i) => conta[l] > 1 ? l + ' (' + ids[i] + ')' : l);
  }

  function titleBlock(ws, titulo, sub, L) {
    const last = colLetter(L.last);
    ws.mergeCells('A1:' + last + '1');
    const t = ws.getCell('A1'); t.value = titulo;
    t.font = { name: FONT, size: 12, bold: true, color: { argb: COL.branco } };
    t.fill = fill(COL.titulo); t.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(1).height = 22;
    ws.mergeCells('A2:' + last + '2');
    const s = ws.getCell('A2'); s.value = sub;
    s.font = { name: FONT, size: 9, italic: true, color: { argb: COL.titulo } };
    s.alignment = { horizontal: 'center' };
    // Com muitas empresas a lista de nomes não cabe em uma linha só.
    const largura = 72 + (L.last - 1) * 16.5;
    if (String(sub).length > largura) {
      s.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      ws.getRow(2).height = 12 + Math.ceil(String(sub).length / largura) * 11;
    }
  }

  function colHeaders(ws, rowNum, labels) {
    const r = ws.getRow(rowNum);
    labels.forEach((lab, i) => {
      const c = r.getCell(i + 1); c.value = lab;
      c.font = { name: FONT, size: 10, bold: true, color: { argb: COL.branco } };
      c.fill = fill(COL.titulo);
      c.alignment = { horizontal: i === 0 ? 'left' : 'right', vertical: 'middle' };
      c.border = borderAll();
    });
    r.height = 18;
  }

  // estilo por tipo de linha
  function styleDataRow(ws, rowNum, kind, L) {
    const r = ws.getRow(rowNum);
    let bg = null, bold = false, italic = false, txt = null;
    if (kind === 'root') { bg = COL.raiz; bold = true; txt = COL.titulo; }
    else if (kind === 'section') { bg = COL.section; bold = true; }
    else if (kind === 'subtotal') { bg = COL.total; bold = true; }
    else if (kind === 'grandtotal') { bg = COL.grand; bold = true; txt = COL.titulo; }
    else if (kind === 'total') { bg = COL.total; bold = true; }
    else if (kind === 'inter') { bg = COL.inter; italic = true; txt = COL.txtInter; }
    else if (kind === 'alerta') { bg = COL.alerta; bold = true; txt = COL.txtAlerta; }
    for (let c = 1; c <= L.last; c++) {
      const cell = r.getCell(c);
      cell.font = { name: FONT, size: 10, bold, italic, color: txt ? { argb: txt } : undefined };
      if (bg) cell.fill = fill(bg);
      cell.border = borderAll();
      if (c === 1) cell.alignment = { wrapText: true, vertical: 'top' };
      else { cell.alignment = { horizontal: 'right' }; cell.numFmt = NUMFMT; }
    }
  }

  // Escreve uma linha de dados. Soma = soma das colunas de empresa e
  // Consolidado = Soma + Eliminações, ambas por fórmula (recalculáveis no Excel).
  // Se line.comps existir (índices das linhas componentes) e `base` for a
  // primeira linha de dados, cada coluna de empresa sai como fórmula de soma
  // dos seus componentes.
  function writeRow(ws, rowNum, line, base, L) {
    const row = ws.getRow(rowNum);
    row.getCell(1).value = line.desc;
    if (line.header) { styleDataRow(ws, rowNum, line.kind, L); return; }
    if (line.kind === 'alerta') {
      row.getCell(L.cons).value = round2(line.consolidado);
      styleDataRow(ws, rowNum, line.kind, L); return;
    }
    const vals = line.vals || [];
    const usaComps = line.comps && line.comps.length && base !== undefined;
    const somaTermos = [];
    for (let i = 0; i < L.n; i++) {
      const nCol = L.first + i, letra = colLetter(nCol);
      if (usaComps) {
        row.getCell(nCol).value = {
          formula: line.comps.map(k => letra + (base + k)).join('+'),
          result: round2(vals[i])
        };
      } else {
        row.getCell(nCol).value = round2(vals[i]);
      }
      somaTermos.push(letra + rowNum);
    }
    row.getCell(L.soma).value = { formula: somaTermos.join('+'), result: round2(line.soma) };
    row.getCell(L.elim).value = line.elim ? round2(line.elim) : 0;
    row.getCell(L.cons).value = {
      formula: colLetter(L.soma) + rowNum + '+' + colLetter(L.elim) + rowNum,
      result: round2(line.consolidado)
    };
    styleDataRow(ws, rowNum, line.kind, L);
    const descLen = String(line.desc || '').length;
    if (descLen > 68) ws.getRow(rowNum).height = descLen > 125 ? 42 : 29;
  }

  function notesCell(ws, rowNum, text, color, L) {
    const last = colLetter(L.last);
    ws.mergeCells('A' + rowNum + ':' + last + rowNum);
    const c = ws.getCell('A' + rowNum); c.value = text;
    c.font = { name: FONT, size: 9, color: { argb: color || 'FF333333' } };
    c.alignment = { wrapText: true, vertical: 'top' };
    ws.getRow(rowNum).height = Math.min(220, 14 + Math.ceil(text.length / 95) * 13);
  }

  function setupCols(ws, L) {
    ws.getColumn(1).width = 72;
    for (let c = 2; c <= L.last; c++) ws.getColumn(c).width = 16.5;
  }

  // ----- Balanço -----
  function sheetBalanco(wb, emps, cons, decisions) {
    const ws = wb.addWorksheet('Balanço Consolidado', { views: [{ showGridLines: false }] });
    const L = layoutFor(emps.length);
    setupCols(ws, L);
    const hs = emps.map(e => e.balanco.header);
    const data = hs.map(h => h.data).find(Boolean) || '';
    const labels = companyLabels(hs.map(h => h.nome), emps.map(e => e.id));
    titleBlock(ws, 'BALANÇO PATRIMONIAL CONSOLIDADO — Saldos atuais em ' + data,
      hs.map(h => h.nome).join(' + ') + ' — empresas independentes', L);
    const hRow = 3;
    colHeaders(ws, hRow, ['Descrição'].concat(labels.map(l => l + ' (R$)'))
      .concat(['Soma (R$)', 'Eliminações (R$)', 'Consolidado (R$)']));
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: hRow }];
    let r = hRow + 1;
    cons.rows.forEach(line => { writeRow(ws, r, line, undefined, L); r++; });

    // notas
    r++;
    const rot = id => { const i = emps.findIndex(e => e.id === id); return i >= 0 ? labels[i] : id; };
    const pares = (decisions.pairs || []).filter(p => p.decisao === 'eliminar');
    let nota = 'NOTAS DA CONSOLIDAÇÃO (saldos atuais em ' + data + '):\n' +
      '1) Empresas independentes (sem relação de controle) — consolidação por somatório dos saldos atuais, com eliminação de saldos recíprocos entre elas.\n';
    // Numeração corrida: o número de pares muda conforme as empresas carregadas.
    let num = 2;
    pares.forEach(p => {
      nota += num + ') ELIMINAÇÃO INTERCOMPANY ' + rot(p.a.company) + ' ↔ ' + rot(p.b.company) +
        ': recíproco confirmável de R$ ' + U.fmt(p.confirmavel) + ' eliminado em ambos os lados (' + p.tratamento + ').\n';
      num++;
      if (p.divergencia > 0.02) {
        nota += num + ') DIVERGÊNCIA: as duas pontas NÃO se conciliam (diferença de R$ ' + U.fmt(p.divergencia) +
          '). O saldo remanescente permanece registrado e destacado como "Saldo intercompany residual não conciliado". Isso indica ERRO DE ESCRITURAÇÃO na origem — recomenda-se conciliar as contas entre as empresas e corrigir o lançamento.\n';
        num++;
      }
    });
    nota += num + ') Valores fiéis aos arquivos anexados; subtotais e totais são calculados por fórmula e recalculáveis.';
    notesCell(ws, r, nota, null, L);
    return ws;
  }

  // ----- DRE -----
  function sheetDRE(wb, emps, dcons) {
    const ws = wb.addWorksheet('DRE Consolidada', { views: [{ showGridLines: false }] });
    const L = layoutFor(emps.length);
    setupCols(ws, L);
    const hs = emps.map(e => e.dre.header);
    const data = hs.map(h => h.data).find(Boolean) || '';
    const labels = companyLabels(hs.map(h => h.nome), emps.map(e => e.id));
    titleBlock(ws, 'DRE CONSOLIDADA — Saldos atuais do período encerrado em ' + data,
      hs.map(h => h.nome).join(' + '), L);
    const hRow = 3;
    colHeaders(ws, hRow, ['Descrição'].concat(labels.map(l => l + ' (R$)'))
      .concat(['Soma (R$)', 'Eliminações (R$)', 'Consolidado (R$)']));
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: hRow }];
    const base = hRow + 1;
    let r = base;
    dcons.lines.forEach(line => { writeRow(ws, r, line, base, L); r++; });
    r++;
    notesCell(ws, r,
      'NOTAS: 1) Não foram identificadas operações de compra/venda entre as empresas no período — não há eliminação de receitas/custos intercompany na DRE; apenas a soma dos saldos atuais. ' +
      '2) Resultado Líquido consolidado recalculado pelas contas da DRE: ' +
      (dcons.resLiq || []).map(v => U.fmt(v)).join(' + ') + ' = ' + U.fmt(dcons.resultLiquido) + '. ' +
      '3) Caso futuramente se identifique faturamento entre as empresas, as receitas/custos correspondentes deverão ser eliminados na coluna de Eliminações.',
      null, L);
    return ws;
  }

  // ----- Conferência -----
  function sheetConferencia(wb, emps, cons, dcons, decisions, fileNames) {
    const ws = wb.addWorksheet('Conferência e Validação', { views: [{ showGridLines: false }] });
    ws.getColumn(1).width = 46; ws.getColumn(2).width = 20; ws.getColumn(3).width = 20; ws.getColumn(4).width = 22;
    let r = 1;
    const title = txt => { ws.mergeCells(`A${r}:D${r}`); const c = ws.getCell(`A${r}`); c.value = txt;
      c.font = { name: FONT, size: 12, bold: true, color: { argb: COL.branco } }; c.fill = fill(COL.titulo);
      c.alignment = { horizontal: 'center' }; ws.getRow(r).height = 20; r++; };
    const line = (a, b, c, d, o) => { o = o || {}; const row = ws.getRow(r);
      row.getCell(1).value = a; if (b !== undefined) row.getCell(2).value = b;
      if (c !== undefined) row.getCell(3).value = c; if (d !== undefined) row.getCell(4).value = d;
      for (let i = 1; i <= 4; i++) { const cell = row.getCell(i);
        cell.font = { name: FONT, size: 10, bold: !!o.bold, color: o.txt ? { argb: o.txt } : undefined };
        if (o.bg) cell.fill = fill(o.bg);
        if (i > 1) { cell.alignment = { horizontal: 'right' }; if (typeof cell.value === 'number') cell.numFmt = NUMFMT; } }
      r++; };

    title('CONFERÊNCIA E VALIDAÇÃO');
    line('Gerado em: ' + new Date().toLocaleString('pt-BR'));
    emps.forEach(e => line('Empresa ' + e.id + ': ' + e.balanco.header.nome +
      ' (CNPJ ' + (e.balanco.header.cnpjFmt || '-') + ')'));
    if (fileNames) line('Arquivos: ' + fileNames.join(' | '));
    r++;
    title('1) Balanço fecha?');
    line('Total do Ativo consolidado', cons.totalAtivo, '', '', { bold: true });
    line('Total do Passivo + PL consolidado', cons.totalPassivo, '', '', { bold: true });
    line('Diferença', cons.diff, '', '', { bold: true, bg: cons.fecha ? COL.total : COL.inter });
    line(cons.fecha ? '✓ BALANÇO FECHA' : '✗ NÃO FECHA — bloqueado', '', '', '', { bold: true, txt: cons.fecha ? COL.titulo : COL.txtInter });
    r++;
    title('2) Resultado Líquido = ' + emps.map(e => e.id).join(' + '));
    emps.forEach((e, ci) => line('Resultado Líquido Empresa ' + e.id, (dcons.resLiq || [])[ci]));
    line('Resultado Líquido Consolidado', dcons.resultLiquido, '', '', { bold: true, bg: dcons.bate ? COL.total : COL.inter });
    r++;
    title('3) Eliminações intercompany');
    line('Par (conta ↔ contraparte)', 'Valor', 'Valor', 'Decisão', { bold: true, bg: COL.section });
    (decisions.pairs || []).forEach(p => {
      line('[' + p.a.company + '] ' + short(p.a.desc) + ' ↔ [' + p.b.company + '] ' + short(p.b.desc),
        p.valA, p.valB, p.decisao === 'eliminar' ? ('eliminar/' + p.tratamento) : 'manter');
      if (p.divergencia > 0.02) line('   → divergência entre pontas', p.divergencia, '', '', { bg: COL.alerta, txt: COL.txtAlerta, bold: true });
    });
    if (!(decisions.pairs || []).length) line('Nenhuma eliminação aplicada.');
    r++;

    // 4) Auditoria: sintética == soma das analíticas/filhos diretos (≤ R$ 0,02)
    title('4) Auditoria — sintéticas × soma das analíticas (Balanço)');
    line('Conta sintética', 'Saldo sintético', 'Soma dos filhos', 'Diferença', { bold: true, bg: COL.section });
    let balDiv = 0;
    emps.forEach(emp => {
      (global.C.auditBalance(emp) || []).forEach(it => {
        if (it.ok) return;
        balDiv++;
        line('[' + emp.id + '] ' + short(it.conta), round2(it.sintetico), round2(it.soma), round2(it.dif), { bg: COL.alerta, txt: COL.txtAlerta });
      });
    });
    if (!balDiv) line('OK — todas as sintéticas conferem com a soma das analíticas.', '', '', '', { bold: true, txt: COL.titulo });
    r++;

    title('5) Auditoria — DRE (blocos e subtotais oficiais)');
    line('Verificação', 'Recalculado', 'Oficial/Soma', 'Diferença', { bold: true, bg: COL.section });
    let dreDiv = 0;
    (dcons.blockAudit || []).forEach(it => {
      if (it.ok) return;
      dreDiv++;
      line('[' + it.empresa + '] ' + short(it.conta), round2(it.sintetico), round2(it.soma), round2(it.dif), { bg: COL.alerta, txt: COL.txtAlerta });
    });
    (dcons.checks || []).forEach(c => {
      line('[' + c.empresa + '] ' + c.conta, round2(c.calc), round2(c.oficial), round2(c.dif),
        c.ok ? {} : { bg: COL.alerta, txt: COL.txtAlerta, bold: true });
      if (!c.ok) dreDiv++;
    });
    if (!dreDiv) line('OK — todos os subtotais da DRE conferem (tolerância ≤ R$ 0,02).', '', '', '', { bold: true, txt: COL.titulo });
    return ws;
  }

  function short(s) { s = String(s || '').trim(); return s.length > 38 ? s.slice(0, 36) + '…' : s; }

  async function generate(emps, cons, dcons, decisions, fileNames) {
    const wb = new global.ExcelJS.Workbook();
    wb.creator = 'Consolidador'; wb.created = new Date();
    sheetBalanco(wb, emps, cons, decisions);
    sheetDRE(wb, emps, dcons);
    sheetConferencia(wb, emps, cons, dcons, decisions, fileNames);
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  global.X = { generate };
})(typeof window !== 'undefined' ? window : globalThis);
