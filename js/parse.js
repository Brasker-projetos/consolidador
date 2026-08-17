/* parse.js — leitura dos .xls/.xlsx (SheetJS) e extração estruturada. */
(function (global) {
  'use strict';
  const U = global.U;

  // Lê o primeiro sheet como matriz de linhas (array de arrays), preservando espaços.
  function sheetToGrid(workbook) {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    // raw:true mantém números; defval:'' evita buracos; blankrows mantidas p/ indentação.
    return global.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
  }

  // Procura em todas as células um rótulo "Chave:" e devolve o texto à direita.
  function findField(grid, regex) {
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] == null ? '' : row[c]);
        const m = cell.match(regex);
        if (m) {
          if (m[1] && m[1].trim()) return m[1].trim();
          // valor pode estar na próxima célula não vazia da mesma linha
          for (let k = c + 1; k < row.length; k++) {
            const nx = String(row[k] == null ? '' : row[k]).trim();
            if (nx) return nx;
          }
        }
      }
    }
    return '';
  }

  // Cabeçalho comum (empresa, cnpj, data).
  function parseHeader(grid) {
    const nome = findField(grid, /empresa\s*:?\s*(.*)/i);
    const cnpj = findField(grid, /c\.?\s*n\.?\s*p\.?\s*j\.?\s*:?\s*(.*)/i);
    let data = findField(grid, /(?:encerrad[oa]\s+em|exerc[ií]cio\s+em|em)\s*:?\s*([0-3]?\d\/[01]?\d\/\d{4})/i);
    if (!data) {
      // procura qualquer dd/mm/aaaa no topo
      for (let r = 0; r < Math.min(grid.length, 15); r++) {
        const row = grid[r] || [];
        for (const cell of row) {
          const m = String(cell == null ? '' : cell).match(/([0-3]?\d\/[01]?\d\/\d{4})/);
          if (m) { data = m[1]; break; }
        }
        if (data) break;
      }
    }
    return { nome: nome, cnpj: U.digits(cnpj), cnpjFmt: cnpj, data: data };
  }

  // Localiza a linha de títulos das colunas e mapeia os índices relevantes.
  // Retorna {headerRow, colDesc, colAtual, colAnterior}.
  function findColumns(grid) {
    for (let r = 0; r < grid.length; r++) {
      const row = (grid[r] || []).map(x => U.norm(x));
      let colDesc = -1, colAtual = -1, colAnterior = -1;
      for (let c = 0; c < row.length; c++) {
        const t = row[c];
        if (colDesc < 0 && /(DESCRICAO|HISTORICO|CONTA|CLASSIFICACAO)/.test(t)) colDesc = c;
        if (/SALDO\s*ATUAL/.test(t)) colAtual = c;
        else if (/SALDO\s*ANTERIOR/.test(t)) colAnterior = c;
        else if (colAtual < 0 && /^SALDO$/.test(t)) colAtual = c; // DRE: coluna "Saldo"
      }
      if (colAtual >= 0 && (colDesc >= 0 || colAnterior >= 0)) {
        return { headerRow: r, colDesc: colDesc, colAtual: colAtual, colAnterior: colAnterior };
      }
    }
    return null;
  }

  // Heurística: acha a coluna de descrição (a célula de texto mais "rica" do corpo).
  function guessDescCol(grid, startRow, colAtual) {
    const score = {};
    for (let r = startRow; r < grid.length; r++) {
      const row = grid[r] || [];
      for (let c = 0; c < row.length; c++) {
        if (c === colAtual) continue;
        const v = row[c];
        if (typeof v === 'string' && v.trim().length > 3 && !/^[\d.,()\-\s]+$/.test(v)) {
          score[c] = (score[c] || 0) + 1;
        }
      }
    }
    let best = -1, bestN = -1;
    Object.keys(score).forEach(c => { if (score[c] > bestN) { bestN = score[c]; best = +c; } });
    return best;
  }

  const FOOTER = /(SISTEMA LICENCIAD|FOLHA\s*:|EMISSAO|EMISSÃO|HORA\s*:|CONTADOR|CRC|ASSINATURA|RECONHEC|^CPF|PAGINA|PÁGINA)/i;

  // Extrai as linhas de conta (descrição + saldo atual + saldo anterior + nível).
  function extractRows(grid, cols) {
    let colDesc = cols.colDesc;
    if (colDesc < 0) colDesc = guessDescCol(grid, cols.headerRow + 1, cols.colAtual);
    const rows = [];
    for (let r = cols.headerRow + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      let descRaw = colDesc >= 0 ? row[colDesc] : '';
      if (descRaw == null) descRaw = '';
      // fallback: se a coluna desc veio vazia, pega a 1ª string não numérica da linha
      if (!String(descRaw).trim()) {
        for (let c = 0; c < row.length; c++) {
          if (c === cols.colAtual || c === cols.colAnterior) continue;
          const v = row[c];
          if (typeof v === 'string' && v.trim() && !/^[\d.,()\-\s]+$/.test(v)) { descRaw = v; break; }
        }
      }
      const descStr = String(descRaw);
      const descTrim = descStr.trim();
      if (!descTrim) continue;
      if (FOOTER.test(descTrim)) continue;
      const atual = U.parseNum(row[cols.colAtual]);
      const anterior = cols.colAnterior >= 0 ? U.parseNum(row[cols.colAnterior]) : 0;
      // ignora linhas sem qualquer número se também não parecem cabeçalho de grupo
      const hasNum = String(row[cols.colAtual] || '').trim() !== '' || atual !== 0;
      rows.push({
        descricao: descTrim,
        descRaw: descStr,
        indent: U.leadSpaces(descStr),
        saldoAtual: atual,
        saldoAnterior: anterior,
        hasNum: hasNum,
        rowIndex: r
      });
    }
    return rows;
  }

  // Atribui nível hierárquico. Usa indentação se houver; senão, infere por rótulos.
  function assignLevels(rows) {
    const indents = rows.map(r => r.indent).filter((v, i, a) => true);
    const maxIndent = Math.max.apply(null, rows.map(r => r.indent).concat([0]));
    if (maxIndent >= 3) {
      // Indentação preservada — deriva passo (menor incremento > 0).
      const uniq = Array.from(new Set(rows.map(r => r.indent))).sort((a, b) => a - b);
      let step = 3;
      for (let i = 1; i < uniq.length; i++) { const d = uniq[i] - uniq[i - 1]; if (d > 0) { step = d; break; } }
      rows.forEach(r => { r.level = Math.round(r.indent / step); });
    } else {
      // Sem indentação: nível por rótulos sintéticos conhecidos (balanço).
      rows.forEach(r => { r.level = levelByLabel(r.descricao); });
    }
    // sintética = tem filho com nível maior logo abaixo
    for (let i = 0; i < rows.length; i++) {
      let synth = false;
      for (let j = i + 1; j < rows.length; j++) {
        if (rows[j].level <= rows[i].level) break;
        synth = true; break;
      }
      rows[i].synthetic = synth;
    }
    return rows;
  }

  function levelByLabel(desc) {
    const n = U.norm(desc);
    if (/^(ATIVO|PASSIVO)$/.test(n)) return 0;
    if (/(CIRCULANTE|NAO[\s-]?CIRCULANTE|PATRIMONIO LIQUIDO|REALIZAVEL)/.test(n)) return 1;
    if (/(DISPONIVEL|CLIENTES|ESTOQUE|FORNECEDOR|OBRIGAC|IMOBILIZAD|CAPITAL|LUCRO|PREJUIZO|EMPRESTIMO|MUTUO|ADIANT|CREDITO)/.test(n)) return 2;
    return 3;
  }

  // ---------- BALANÇO ----------
  function parseBalance(grid, fileName) {
    const header = parseHeader(grid);
    let cols = findColumns(grid);
    if (!cols) throw new Error('Não encontrei a linha de títulos (Descrição/Saldo Atual) em ' + fileName);
    let rows = extractRows(grid, cols);
    rows = assignLevels(rows);

    // Apara lixo: mantém do ATIVO até o fim do bloco PASSIVO (descarta DRE anexa, assinaturas).
    const ativoIdx = rows.findIndex(r => r.level === 0 && /^ATIVO$/.test(U.norm(r.descricao)));
    const passivoIdx = rows.findIndex(r => r.level === 0 && /^PASSIVO$/.test(U.norm(r.descricao)));
    if (ativoIdx >= 0 && passivoIdx > ativoIdx) {
      let end = rows.length;
      for (let i = passivoIdx + 1; i < rows.length; i++) { if (rows[i].level === 0) { end = i; break; } }
      rows = rows.slice(ativoIdx, end);
    }

    // Totais e validação de fechamento.
    const tAtivo = findTotal(rows, /^ATIVO$/);
    const tPassivo = findTotal(rows, /^PASSIVO$/);
    let fecha = null;
    if (tAtivo != null && tPassivo != null) {
      fecha = Math.abs(Math.abs(tAtivo) - Math.abs(tPassivo)) <= 0.02;
    }
    return {
      tipo: 'balanco', fileName, header, rows,
      totalAtivo: tAtivo, totalPassivo: tPassivo, fecha
    };
  }

  function findTotal(rows, regex) {
    for (const r of rows) { if (regex.test(U.norm(r.descricao))) return r.saldoAtual; }
    return null;
  }

  // ---------- DRE ----------
  // Linha é subtotal/total (negrito) quando casa um destes rótulos.
  const DRE_SUBTOTAL = /^(RECEITA LIQUIDA|LUCRO BRUTO|RESULTADO OPERACIONAL|RESULTADO ANTES|RESULTADO LIQUIDO|LUCRO LIQUIDO|PREJUIZO|RESULTADO DO EXERC)/;
  const DRE_RESULTADO = /(RESULTADO LIQUIDO|LUCRO LIQUIDO|PREJUIZO|RESULTADO DO EXERC)/;
  // Atenção: "FOLHA"/"HORA" só com dois-pontos (rodapé) — existem contas
  // legítimas como "FOLHA DE PAGAMENTO" e "HORAS EXTRAS".
  const DRE_SKIP = /(CPF|CRC|SISTEMA LICENC|DEMONSTRACAO DO RESULTADO|SOCIO|ADMINISTRADOR|REG\. NO|^FOLHA\s*:|EMISSAO|^HORA\s*:|NUMERO LIVRO)/;

  // Extrai as linhas da DRE. A hierarquia vem da COLUNA da descrição
  // (grupos em uma coluna à esquerda, detalhes em colunas à direita), não de espaços.
  // A coluna estrutural é RELATIVA: a menor coluna de descrição usada no documento —
  // alguns exports contábeis recuam tudo (grupos não ficam na coluna 0).
  function extractDRE(grid, cols) {
    const lines = [];
    for (let r = cols.headerRow + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const rawSaldo = row[cols.colAtual];
      if (rawSaldo === '' || rawSaldo === null || rawSaldo === undefined) continue; // pula título/assinatura
      // Descrição = ÚLTIMA célula textual antes da coluna de saldo, ignorando
      // códigos de classificação (ex.: "4.1.01.020.001") e células numéricas.
      let descCol = -1, desc = '';
      for (let c = 0; c < cols.colAtual; c++) {
        const v = row[c];
        if (typeof v === 'string' && v.trim() && !/^[\d.,()\-\s]+$/.test(v.trim())) { descCol = c; desc = v.trim(); }
      }
      if (descCol < 0) continue;
      const n = U.norm(desc);
      if (DRE_SKIP.test(n)) continue;
      lines.push({
        descricao: desc, normLabel: n, descCol,
        level: 1, // ajustado abaixo de forma relativa
        saldoAtual: U.parseNum(rawSaldo),
        isSubtotal: DRE_SUBTOTAL.test(n),
        rowIndex: r
      });
    }
    // Nível relativo: linhas na menor coluna de descrição são estruturais (nível 0).
    const minCol = lines.reduce((m, l) => Math.min(m, l.descCol), Infinity);
    lines.forEach(l => { l.level = l.descCol === minCol ? 0 : 1; });
    return lines;
  }

  function parseDRE(grid, fileName) {
    const header = parseHeader(grid);
    let cols = findColumns(grid);
    if (!cols) throw new Error('Não encontrei colunas de Saldo na DRE ' + fileName);
    const lines = extractDRE(grid, cols);

    // Blocos: cada sintética (menor coluna) com suas analíticas (colunas à direita).
    const blocks = []; let cur = null;
    lines.forEach(l => {
      const key = U.dreKey(l.descricao);
      if (l.level === 0) { cur = { descricao: l.descricao, key, value: l.saldoAtual, analytics: [] }; blocks.push(cur); }
      else if (cur) cur.analytics.push({ descricao: l.descricao, key, value: l.saldoAtual });
      else blocks.push({ descricao: l.descricao, key, value: l.saldoAtual, analytics: [] });
    });

    // Linhas calculadas (subtotais oficiais do arquivo) — usadas só p/ validação.
    const COMPUTED = /^(RECEITA LIQUIDA|LUCRO BRUTO|RESULTADO OPERACIONAL|RESULTADO ANTES|RESULTADO LIQUIDO|LUCRO LIQUIDO|PREJUIZO|RESULTADO DO EXERC)/;
    const computed = blocks.filter(b => COMPUTED.test(b.key));
    const official = {
      receitaLiquida: (computed.find(b => /^RECEITA LIQUIDA/.test(b.key)) || {}).value,
      // ROL após custos: "LUCRO BRUTO" ou 1ª ocorrência de "RESULTADO OPERACIONAL LÍQUIDO"
      lucroBruto: (computed.find(b => /^(LUCRO BRUTO|RESULTADO OPERACIONAL)/.test(b.key)) || {}).value,
      resultadoOperacional: undefined,
      resultLiquido: undefined
    };
    // ROL após despesas: última ocorrência de RESULTADO OPERACIONAL/ANTES.
    computed.forEach(b => { if (/^(RESULTADO OPERACIONAL|RESULTADO ANTES)/.test(b.key)) official.resultadoOperacional = b.value; });

    // Resultado líquido = última linha calculada de resultado/lucro/prejuízo.
    // Ausência do total oficial não significa lucro/prejuízo zero. Mantemos
    // `undefined` para que a validação informe a limitação de leitura, sem
    // criar uma divergência artificial contra o resultado recalculado.
    let resultLiquido = undefined, found = false;
    for (let i = computed.length - 1; i >= 0; i--) {
      if (DRE_RESULTADO.test(computed[i].key)) { resultLiquido = computed[i].value; found = true; break; }
    }
    if (!found && computed.length) resultLiquido = computed[computed.length - 1].value;
    official.resultLiquido = resultLiquido;

    // Sanidade de parsing (não rígida): identidades que devem fechar nos arquivos.
    const sanity = [];
    const rb = (blocks.find(b => /^RECEITAS? BRUTA/.test(b.key)) || {}).value;
    const ded = (blocks.find(b => /^DEDUC/.test(b.key)) || {}).value;
    if (official.receitaLiquida !== undefined && rb !== undefined && ded !== undefined
        && Math.abs(official.receitaLiquida - (rb + ded)) > 0.02)
      sanity.push('Receita Líquida não fecha (RB+Deduções)');
    if (resultLiquido === undefined)
      sanity.push('Resultado Líquido oficial não identificado no arquivo');

    return { tipo: 'dre', fileName, header, lines, blocks, official, resultLiquido, sanity };
  }

  global.P = { parseBalance, parseDRE, sheetToGrid, parseHeader };
})(typeof window !== 'undefined' ? window : globalThis);
