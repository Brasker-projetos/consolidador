/* xlsxout.js — gera o .xlsx consolidado no layout da referência (ExcelJS). */
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
  function companyLabels(a, b) {
    const tokens = s => U.norm(s).split(/\s+/).filter(x => x && !/^(LTDA|EIRELI|S\/A|SA)$/.test(x));
    const ta = tokens(a), tb = tokens(b), setA = new Set(ta), setB = new Set(tb);
    const la = ta.find(x => !setB.has(x)) || firstWord(a);
    const lb = tb.find(x => !setA.has(x)) || firstWord(b);
    return [la, lb];
  }

  function titleBlock(ws, titulo, sub) {
    ws.mergeCells('A1:F1');
    const t = ws.getCell('A1'); t.value = titulo;
    t.font = { name: FONT, size: 12, bold: true, color: { argb: COL.branco } };
    t.fill = fill(COL.titulo); t.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(1).height = 22;
    ws.mergeCells('A2:F2');
    const s = ws.getCell('A2'); s.value = sub;
    s.font = { name: FONT, size: 9, italic: true, color: { argb: COL.titulo } };
    s.alignment = { horizontal: 'center' };
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
  function styleDataRow(ws, rowNum, kind) {
    const r = ws.getRow(rowNum);
    let bg = null, bold = false, italic = false, txt = null;
    if (kind === 'root') { bg = COL.raiz; bold = true; txt = COL.titulo; }
    else if (kind === 'section') { bg = COL.section; bold = true; }
    else if (kind === 'subtotal') { bg = COL.total; bold = true; }
    else if (kind === 'grandtotal') { bg = COL.grand; bold = true; txt = COL.titulo; }
    else if (kind === 'total') { bg = COL.total; bold = true; }
    else if (kind === 'inter') { bg = COL.inter; italic = true; txt = COL.txtInter; }
    else if (kind === 'alerta') { bg = COL.alerta; bold = true; txt = COL.txtAlerta; }
    for (let c = 1; c <= 6; c++) {
      const cell = r.getCell(c);
      cell.font = { name: FONT, size: 10, bold, italic, color: txt ? { argb: txt } : undefined };
      if (bg) cell.fill = fill(bg);
      cell.border = borderAll();
      if (c === 1) cell.alignment = { wrapText: true, vertical: 'top' };
      else { cell.alignment = { horizontal: 'right' }; cell.numFmt = NUMFMT; }
    }
  }

  // escreve uma linha de dados (com fórmulas Soma=B+C e Consolidado=D+E).
  // Se line.comps existir (índices das linhas componentes) e `base` for a
  // primeira linha de dados, B e C saem como fórmula de soma dos componentes.
  function writeRow(ws, rowNum, line, base) {
    const row = ws.getRow(rowNum);
    row.getCell(1).value = line.desc;
    if (line.header) { styleDataRow(ws, rowNum, line.kind); return; }
    if (line.kind === 'alerta') {
      row.getCell(6).value = round2(line.consolidado);
      styleDataRow(ws, rowNum, line.kind); return;
    }
    if (line.comps && line.comps.length && base !== undefined) {
      const f = col => line.comps.map(i => col + (base + i)).join('+');
      row.getCell(2).value = { formula: f('B'), result: round2(line.valA) };
      row.getCell(3).value = { formula: f('C'), result: round2(line.valB) };
    } else {
      row.getCell(2).value = round2(line.valA);
      row.getCell(3).value = round2(line.valB);
    }
    row.getCell(4).value = { formula: `B${rowNum}+C${rowNum}`, result: round2(line.soma) };
    row.getCell(5).value = line.elim ? round2(line.elim) : 0;
    row.getCell(6).value = { formula: `D${rowNum}+E${rowNum}`, result: round2(line.consolidado) };
    styleDataRow(ws, rowNum, line.kind);
    const descLen = String(line.desc || '').length;
    if (descLen > 68) ws.getRow(rowNum).height = descLen > 125 ? 42 : 29;
  }

  function notesCell(ws, rowNum, text, color) {
    ws.mergeCells(`A${rowNum}:F${rowNum}`);
    const c = ws.getCell(`A${rowNum}`); c.value = text;
    c.font = { name: FONT, size: 9, color: { argb: color || 'FF333333' } };
    c.alignment = { wrapText: true, vertical: 'top' };
    ws.getRow(rowNum).height = Math.min(220, 14 + Math.ceil(text.length / 95) * 13);
  }

  function setupCols(ws) {
    ws.getColumn(1).width = 72;
    [2, 3, 4, 5, 6].forEach(c => ws.getColumn(c).width = 16.5);
  }

  // ----- Balanço -----
  function sheetBalanco(wb, empA, empB, cons, decisions) {
    const ws = wb.addWorksheet('Balanço Consolidado', { views: [{ showGridLines: false }] });
    setupCols(ws);
    const hA = empA.balanco.header, hB = empB.balanco.header;
    const data = hA.data || hB.data || '';
    const labels = companyLabels(hA.nome, hB.nome);
    titleBlock(ws, 'BALANÇO PATRIMONIAL CONSOLIDADO — Saldos atuais em ' + data,
      hA.nome + ' + ' + hB.nome + ' — empresas independentes');
    const hRow = 3;
    colHeaders(ws, hRow, ['Descrição', labels[0] + ' (R$)', labels[1] + ' (R$)',
      'Soma (R$)', 'Eliminações (R$)', 'Consolidado (R$)']);
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: hRow }];
    let r = hRow + 1;
    cons.rows.forEach(line => { writeRow(ws, r, line); r++; });

    // notas
    r++;
    const pares = (decisions.pairs || []).filter(p => p.decisao === 'eliminar');
    let nota = 'NOTAS DA CONSOLIDAÇÃO (saldos atuais em ' + data + '):\n' +
      '1) Empresas independentes (sem relação de controle) — consolidação por somatório dos saldos atuais, com eliminação de saldos recíprocos entre elas.\n';
    pares.forEach(p => {
      nota += '2) ELIMINAÇÃO INTERCOMPANY: recíproco confirmável de R$ ' + U.fmt(p.confirmavel) + ' eliminado em ambos os lados (' + p.tratamento + ').\n';
      if (p.divergencia > 0.02)
        nota += '3) DIVERGÊNCIA: as duas pontas NÃO se conciliam (diferença de R$ ' + U.fmt(p.divergencia) + '). O saldo remanescente permanece registrado e destacado como "Saldo intercompany residual não conciliado". Isso indica ERRO DE ESCRITURAÇÃO na origem — recomenda-se conciliar as contas entre as empresas e corrigir o lançamento.\n';
    });
    nota += '4) Valores fiéis aos arquivos anexados; subtotais e totais são calculados por fórmula e recalculáveis.';
    notesCell(ws, r, nota);
    return ws;
  }

  // ----- DRE -----
  function sheetDRE(wb, empA, empB, dcons) {
    const ws = wb.addWorksheet('DRE Consolidada', { views: [{ showGridLines: false }] });
    setupCols(ws);
    const hA = empA.dre.header, hB = empB.dre.header, data = hA.data || hB.data || '';
    const labels = companyLabels(hA.nome, hB.nome);
    titleBlock(ws, 'DRE CONSOLIDADA — Saldos atuais do período encerrado em ' + data,
      hA.nome + ' + ' + hB.nome);
    const hRow = 3;
    colHeaders(ws, hRow, ['Descrição', labels[0] + ' (R$)', labels[1] + ' (R$)',
      'Soma (R$)', 'Eliminações (R$)', 'Consolidado (R$)']);
    ws.views = [{ showGridLines: false, state: 'frozen', ySplit: hRow }];
    const base = hRow + 1;
    let r = base;
    dcons.lines.forEach(line => { writeRow(ws, r, line, base); r++; });
    r++;
    notesCell(ws, r,
      'NOTAS: 1) Não foram identificadas operações de compra/venda entre as empresas no período — não há eliminação de receitas/custos intercompany na DRE; apenas a soma dos saldos atuais. ' +
      '2) Resultado Líquido consolidado recalculado pelas contas da DRE: ' + U.fmt(dcons.resLiqA) + ' + ' + U.fmt(dcons.resLiqB) + ' = ' + U.fmt(dcons.resultLiquido) + '. ' +
      '3) Caso futuramente se identifique faturamento entre as empresas, as receitas/custos correspondentes deverão ser eliminados na coluna de Eliminações.');
    return ws;
  }

  // ----- Conferência -----
  function sheetConferencia(wb, empA, empB, cons, dcons, decisions, fileNames) {
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
    line('Empresa A: ' + empA.balanco.header.nome + ' (CNPJ ' + (empA.balanco.header.cnpjFmt || '-') + ')');
    line('Empresa B: ' + empB.balanco.header.nome + ' (CNPJ ' + (empB.balanco.header.cnpjFmt || '-') + ')');
    if (fileNames) line('Arquivos: ' + fileNames.join(' | '));
    r++;
    title('1) Balanço fecha?');
    line('Total do Ativo consolidado', cons.totalAtivo, '', '', { bold: true });
    line('Total do Passivo + PL consolidado', cons.totalPassivo, '', '', { bold: true });
    line('Diferença', cons.diff, '', '', { bold: true, bg: cons.fecha ? COL.total : COL.inter });
    line(cons.fecha ? '✓ BALANÇO FECHA' : '✗ NÃO FECHA — bloqueado', '', '', '', { bold: true, txt: cons.fecha ? COL.titulo : COL.txtInter });
    r++;
    title('2) Resultado Líquido = A + B');
    line('Resultado Líquido Empresa A', dcons.resLiqA);
    line('Resultado Líquido Empresa B', dcons.resLiqB);
    line('Resultado Líquido Consolidado', dcons.resultLiquido, '', '', { bold: true, bg: dcons.bate ? COL.total : COL.inter });
    r++;
    title('3) Eliminações intercompany');
    line('Par (conta A ↔ conta B)', 'Valor A', 'Valor B', 'Decisão', { bold: true, bg: COL.section });
    (decisions.pairs || []).forEach(p => {
      line(short(p.a.desc) + ' ↔ ' + short(p.b.desc), p.valA, p.valB, p.decisao === 'eliminar' ? ('eliminar/' + p.tratamento) : 'manter');
      if (p.divergencia > 0.02) line('   → divergência entre pontas', p.divergencia, '', '', { bg: COL.alerta, txt: COL.txtAlerta, bold: true });
    });
    if (!(decisions.pairs || []).length) line('Nenhuma eliminação aplicada.');
    r++;

    // 4) Auditoria: sintética == soma das analíticas/filhos diretos (≤ R$ 0,02)
    title('4) Auditoria — sintéticas × soma das analíticas (Balanço)');
    line('Conta sintética', 'Saldo sintético', 'Soma dos filhos', 'Diferença', { bold: true, bg: COL.section });
    let balDiv = 0;
    [['A', empA], ['B', empB]].forEach(([rot, emp]) => {
      (global.C.auditBalance(emp) || []).forEach(it => {
        if (it.ok) return;
        balDiv++;
        line('[' + rot + '] ' + short(it.conta), round2(it.sintetico), round2(it.soma), round2(it.dif), { bg: COL.alerta, txt: COL.txtAlerta });
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

  async function generate(empA, empB, cons, dcons, decisions, fileNames) {
    const wb = new global.ExcelJS.Workbook();
    wb.creator = 'Consolidador'; wb.created = new Date();
    sheetBalanco(wb, empA, empB, cons, decisions);
    sheetDRE(wb, empA, empB, dcons);
    sheetConferencia(wb, empA, empB, cons, dcons, decisions, fileNames);
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  global.X = { generate };
})(typeof window !== 'undefined' ? window : globalThis);
