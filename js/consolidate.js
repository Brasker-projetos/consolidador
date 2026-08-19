/* consolidate.js — intercompany (detecção/divergência), consolidação do balanço
   por seções (AC/ANC/PC/PNC/PL) e DRE gerencial (§3.2). */
(function (global) {
  'use strict';
  const U = global.U;

  const NAME_THRESHOLD = 0.6;

  // Limite de empresas por consolidação. Acima disso o relatório deixa de caber
  // numa leitura confortável (uma coluna por empresa) e a conferência manual,
  // que é o controle final do contador, fica inviável.
  const MAX_EMPRESAS = 10;
  const MIN_EMPRESAS = 2;

  // Toda entrada pública valida a lista antes de calcular: erro cedo e com
  // texto claro é melhor do que planilha consolidada errada.
  function checkEmps(emps) {
    if (!Array.isArray(emps) || !emps.length) throw new Error('Nenhuma empresa recebida para consolidar.');
    if (emps.length < MIN_EMPRESAS) throw new Error('São necessárias pelo menos ' + MIN_EMPRESAS + ' empresas para consolidar.');
    if (emps.length > MAX_EMPRESAS) throw new Error('Limite de ' + MAX_EMPRESAS + ' empresas por consolidação.');
    const ids = new Set();
    emps.forEach((e, i) => {
      if (!e || !e.balanco || !e.dre) throw new Error('Empresa ' + ((e && e.id) || (i + 1)) + ': balanço ou DRE não carregados.');
      if (ids.has(e.id)) throw new Error('Identificador de empresa repetido: ' + e.id + '.');
      ids.add(e.id);
    });
    return emps;
  }

  function sideOf(desc) {
    const n = U.norm(desc);
    if (/(A RECEBER|RECEBER|CLIENTE|CREDITO|MUTUOS?)/.test(n)) return 'receber';
    if (/(A PAGAR|PAGAR|FORNECEDOR|EMPRESTIMO|OBRIGAC)/.test(n)) return 'pagar';
    return 'indef';
  }

  function analytics(emp) {
    return (emp.balanco.rows || []).filter(r => !r.synthetic && r.saldoAtual !== 0)
      .map(r => Object.assign({ rowId: emp.id + ':' + r.rowIndex }, r));
  }

  // ---------- Posições financeiras (clientes, fornecedores e bancos) ----------
  // O balanço traz o saldo por conta, não os títulos com vencimento. Por isso as
  // posições abaixo exibem somente o detalhamento analítico que existe na origem.
  const POSITION_TYPES = {
    clientes: {
      title: 'CLIENTES EM ABERTO', root: 'A',
      match: /\b(CLIENTES?|DUPLICATAS? A RECEBER|CONTAS? A RECEBER|TITULOS? A RECEBER|CREDITOS? DE CLIENTES?)\b/
    },
    fornecedores: {
      title: 'FORNECEDORES EM ABERTO', root: 'P',
      match: /\b(FORNECEDORES?|DUPLICATAS? A PAGAR|CONTAS? A PAGAR)\b/
    },
    bancos: {
      title: 'BANCOS E APLICAÇÕES DE LIQUIDEZ IMEDIATA', root: 'A',
      // Não classifica empréstimos/financiamentos bancários: estes pertencem ao passivo.
      match: /\b(BANCOS?|CONTA CORRENTE|APLICACOES? FINANCEIRAS?)\b/
    }
  };

  function positionRows(emp, type) {
    const rows = emp.balanco.rows || [];
    const out = [];
    let root = null;
    const matchedAtLevel = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i], norm = U.norm(row.descricao);
      if (row.level === 0) {
        root = /PASSIVO/.test(norm) ? 'P' : (/ATIVO/.test(norm) ? 'A' : null);
        matchedAtLevel.length = 0;
        continue;
      }
      // Descarta marcadores que já não são ancestrais do item atual.
      while (matchedAtLevel.length && matchedAtLevel[matchedAtLevel.length - 1] >= row.level) matchedAtLevel.pop();
      if (root !== type.root) continue;
      if (type.match.test(norm)) matchedAtLevel.push(row.level);
      // Somente contas analíticas: evita duplicar o saldo do grupo e dos filhos.
      if (!row.synthetic && matchedAtLevel.length && Math.abs(row.saldoAtual) > 0.000001) {
        out.push({
          rowId: emp.id + ':' + row.rowIndex,
          descricao: row.descricao,
          value: type.root === 'P' ? -row.saldoAtual : row.saldoAtual
        });
      }
    }
    return out;
  }

  function positionEliminations(pairs) {
    const out = new Map();
    (pairs || []).forEach(p => {
      if (p.decisao !== 'eliminar' || p.tratamento === 'manter') return;
      const aAmt = p.tratamento === 'cheio' ? Math.abs(p.valA) : p.confirmavel;
      const bAmt = p.tratamento === 'cheio' ? Math.abs(p.valB) : p.confirmavel;
      out.set(p.a.rowId, (out.get(p.a.rowId) || 0) - aAmt);
      out.set(p.b.rowId, (out.get(p.b.rowId) || 0) - bAmt);
    });
    return out;
  }

  function financialPositions(emps, pairs) {
    checkEmps(emps);
    const eliminations = positionEliminations(pairs);
    const n = emps.length;
    const positions = {};
    Object.keys(POSITION_TYPES).forEach(key => {
      const type = POSITION_TYPES[key], byDesc = new Map();
      const add = (item, ci) => {
        const canon = U.norm(item.descricao);
        if (!byDesc.has(canon)) byDesc.set(canon, { desc: item.descricao, vals: new Array(n).fill(0), elim: 0 });
        const line = byDesc.get(canon);
        line.vals[ci] += item.value;
        line.elim += eliminations.get(item.rowId) || 0;
      };
      emps.forEach((emp, ci) => positionRows(emp, type).forEach(item => add(item, ci)));
      const lines = Array.from(byDesc.values()).map(line => {
        line.soma = line.vals.reduce((s, x) => s + x, 0);
        line.consolidado = line.soma + line.elim;
        return line;
      }).sort((a, b) => a.desc.localeCompare(b.desc, 'pt-BR'));
      positions[key] = {
        title: type.title, lines,
        totais: emps.map((e, ci) => lines.reduce((s, x) => s + x.vals[ci], 0)),
        totalElim: lines.reduce((s, x) => s + x.elim, 0),
        total: lines.reduce((s, x) => s + x.consolidado, 0)
      };
    });
    return positions;
  }

  // ---------- Detecção intercompany ----------
  // Varre todas as combinações de duas empresas (A×B, A×C, B×C…). Uma mesma
  // conta analítica só pode entrar em UM par: sem essa trava, uma conta cujo
  // texto lembra o nome de duas empresas seria eliminada duas vezes.
  function detectIntercompany(emps) {
    checkEmps(emps);
    const rowsOf = emps.map(analytics);
    const nomes = emps.map(e => e.balanco.header.nome);
    const cnpjs = emps.map(e => U.digits(e.balanco.header.cnpj));
    const pairs = [], solos = [], soloIds = new Set();
    const usedRow = new Set();

    // Contas da empresa ci que citam o nome ou o CNPJ da empresa cj.
    const candidatos = (ci, cj) => rowsOf[ci]
      .filter(r => U.nameAppears(r.descricao, nomes[cj]) >= NAME_THRESHOLD ||
        (cnpjs[cj] && U.digits(r.descricao).indexOf(cnpjs[cj]) >= 0))
      .map(r => ({ company: emps[ci].id, companyIndex: ci, rowId: r.rowId,
        desc: r.descricao, value: r.saldoAtual, side: sideOf(r.descricao) }));

    for (let i = 0; i < emps.length; i++) {
      for (let j = i + 1; j < emps.length; j++) {
        const candA = candidatos(i, j), candB = candidatos(j, i);
        const usedA = new Set(), usedB = new Set();
        candA.forEach((a, ia) => {
          if (usedRow.has(a.rowId)) { usedA.add(ia); return; }
          let best = -1, bestScore = -1;
          candB.forEach((b, ib) => {
            if (usedB.has(ib) || usedRow.has(b.rowId)) return;
            let sc = 1;
            if ((a.side === 'receber' && b.side === 'pagar') || (a.side === 'pagar' && b.side === 'receber')) sc += 2;
            if (sc > bestScore) { bestScore = sc; best = ib; }
          });
          if (best >= 0) {
            usedA.add(ia); usedB.add(best);
            const b = candB[best];
            usedRow.add(a.rowId); usedRow.add(b.rowId);
            pairs.push({ a, b, valA: a.value, valB: b.value,
              confirmavel: Math.min(Math.abs(a.value), Math.abs(b.value)),
              divergencia: Math.abs(Math.abs(a.value) - Math.abs(b.value)),
              decisao: 'eliminar', tratamento: 'residual' });
          }
        });
        const solo = c => { if (!usedRow.has(c.rowId) && !soloIds.has(c.rowId)) { soloIds.add(c.rowId); solos.push(c); } };
        candA.forEach((a, ia) => { if (!usedA.has(ia)) solo(a); });
        candB.forEach((b, ib) => { if (!usedB.has(ib)) solo(b); });
      }
    }
    return { pairs, solos };
  }

  // ---------- Modelo de seções do balanço ----------
  const SEC_ORDER = ['AC', 'ANC', 'PC', 'PNC', 'PL'];
  const SEC_NAME = { AC: 'ATIVO CIRCULANTE', ANC: 'ATIVO NÃO-CIRCULANTE',
    PC: 'PASSIVO CIRCULANTE', PNC: 'PASSIVO NÃO-CIRCULANTE', PL: 'PATRIMÔNIO LÍQUIDO' };

  function classL1(norm, root) {
    if (root === 'A') {
      if (/NAO-?CIRCULANTE/.test(norm) || /REALIZAVEL/.test(norm)) return 'ANC';
      if (/CIRCULANTE/.test(norm)) return 'AC';
    } else {
      if (/PATRIMONIO/.test(norm)) return 'PL';
      if (/NAO-?CIRCULANTE/.test(norm)) return 'PNC';
      if (/CIRCULANTE/.test(norm)) return 'PC';
    }
    return null;
  }

  function titleCase(s) {
    return String(s).toLowerCase()
      .replace(/(^|\s|\()([a-zà-ú])/g, (m, p, c) => p + c.toUpperCase())
      .replace(/\b(De|Da|Do|Das|Dos|E|A|O)\b/g, m => m.toLowerCase());
  }

  // ---------- Consolidação do balanço ----------
  function balanceKey(s) {
    return U.norm(s).replace(/[.,;:()]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function treeNodes(rows, start, end, parentLevel, sign, empId) {
    const out = [];
    let i = start;
    while (i < end) {
      const r = rows[i];
      if (r.level !== parentLevel + 1) { i++; continue; }
      let nodeEnd = i + 1;
      while (nodeEnd < end && rows[nodeEnd].level > r.level) nodeEnd++;
      const children = treeNodes(rows, i + 1, nodeEnd, r.level, sign, empId);
      const value = sign * r.saldoAtual;
      // "Em aberto" = saldo atual diferente de zero. Sintéticas com filhos
      // válidos permanecem, mesmo quando seu próprio saldo for zero.
      if (Math.abs(value) > 0.000001 || children.length) {
        out.push({
          label: r.descricao, key: balanceKey(r.descricao), value,
          rowId: empId + ':' + r.rowIndex, children
        });
      }
      i = nodeEnd;
    }
    return out;
  }

  function balanceTree(emp) {
    const rows = emp.balanco.rows || [];
    const sections = {}; SEC_ORDER.forEach(s => { sections[s] = { total: 0, nodes: [] }; });
    const extras = { A: [], P: [] };
    let root = null;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i], n = U.norm(r.descricao);
      if (r.level === 0) { root = /PASSIVO/.test(n) ? 'P' : (/ATIVO/.test(n) ? 'A' : null); continue; }
      if (r.level !== 1 || !root) continue;
      let end = i + 1; while (end < rows.length && rows[end].level > r.level) end++;
      const sec = classL1(n, root);
      const sign = root === 'P' ? -1 : 1;
      const nodes = treeNodes(rows, i + 1, end, r.level, sign, emp.id);
      if (sec) {
        sections[sec].total += sign * r.saldoAtual;
        sections[sec].nodes = sections[sec].nodes.concat(nodes);
      } else if (Math.abs(r.saldoAtual) > 0.000001 || nodes.length) {
        extras[root].push({ label: r.descricao, key: balanceKey(r.descricao),
          value: sign * r.saldoAtual, rowId: emp.id + ':' + r.rowIndex, children: nodes });
      }
      i = end - 1;
    }
    return {
      sections, extras,
      ativo: emp.balanco.totalAtivo || 0,
      passivo: -(emp.balanco.totalPassivo || 0)
    };
  }

  function consolidateBalance(emps, pairs) {
    checkEmps(emps);
    const n = emps.length;
    const trees = emps.map(balanceTree);

    // Eliminação vinculada à conta analítica exata, nunca ao texto genérico
    // da sintética. Isso impede eliminar ou somar contas vizinhas por engano.
    const rowElim = new Map(); let residual = 0;
    (pairs || []).forEach(p => {
      if (p.decisao !== 'eliminar' || p.tratamento === 'manter') return;
      const amtA = p.tratamento === 'cheio' ? Math.abs(p.valA) : p.confirmavel;
      const amtB = p.tratamento === 'cheio' ? Math.abs(p.valB) : p.confirmavel;
      rowElim.set(p.a.rowId, (rowElim.get(p.a.rowId) || 0) - amtA);
      rowElim.set(p.b.rowId, (rowElim.get(p.b.rowId) || 0) - amtB);
      if (p.tratamento !== 'cheio') residual += p.divergencia;
    });

    const out = [];
    // `vals` tem uma posição por empresa, sempre na ordem de `emps`.
    const push = (desc, kind, vals, e) => {
      const v = new Array(n);
      for (let i = 0; i < n; i++) v[i] = vals[i] || 0;
      const soma = v.reduce((s, x) => s + x, 0); const el = e || 0;
      out.push({ desc, kind, vals: v, elim: el, soma, consolidado: soma + el });
    };
    const pushHeader = (desc, kind) => out.push({ desc, kind, header: true });

    function eliminationOf(node) {
      if (!node) return 0;
      return (rowElim.get(node.rowId) || 0) +
        (node.children || []).reduce((s, child) => s + eliminationOf(child), 0);
    }

    // Recebe uma lista de nós POR EMPRESA e emite uma linha por conta, casando
    // pela chave da descrição. As contas da 1ª empresa ditam a ordem; contas
    // exclusivas de uma empresa seguinte entram na ordem original dela, dentro
    // do mesmo pai, antes de o relatório avançar para a próxima sintética.
    function emitTrees(nodeLists, depth) {
      const buckets = nodeLists.map(list => {
        const m = new Map();
        list.forEach((node, i) => {
          if (!m.has(node.key)) m.set(node.key, []);
          m.get(node.key).push({ node, i });
        });
        return m;
      });
      const used = nodeLists.map(() => new Set());
      const emitGroup = group => {
        const childLists = group.map(nd => (nd && nd.children) ? nd.children : []);
        const e = group.reduce((s, nd) => s + eliminationOf(nd), 0);
        const hasChildren = childLists.some(l => l.length);
        push('   '.repeat(depth + 1) + group.find(Boolean).label,
          hasChildren ? 'subtotal' : (Math.abs(e) > 0.000001 ? 'inter' : 'item'),
          group.map(nd => nd ? nd.value : 0), e);
        if (hasChildren) emitTrees(childLists, depth + 1);
      };
      for (let ci = 0; ci < nodeLists.length; ci++) {
        for (let ni = 0; ni < nodeLists[ci].length; ni++) {
          if (used[ci].has(ni)) continue;
          used[ci].add(ni);
          const node = nodeLists[ci][ni];
          const group = new Array(nodeLists.length).fill(null);
          group[ci] = node;
          for (let cj = ci + 1; cj < nodeLists.length; cj++) {
            const bucket = buckets[cj].get(node.key) || [];
            const match = bucket.find(x => !used[cj].has(x.i));
            if (match) { used[cj].add(match.i); group[cj] = match.node; }
          }
          emitGroup(group);
        }
      }
    }

    let ativoElim = 0, passivoElim = 0;
    function emitSection(secKey) {
      pushHeader(SEC_NAME[secKey], 'section');
      const nodeLists = trees.map(t => t.sections[secKey].nodes);
      emitTrees(nodeLists, 0);
      const e = nodeLists.reduce((s, list) => s + list.reduce((s2, x) => s2 + eliminationOf(x), 0), 0);
      push('Total do ' + titleCase(SEC_NAME[secKey]), 'subtotal', trees.map(t => t.sections[secKey].total), e);
      if (secKey === 'AC' || secKey === 'ANC') ativoElim += e; else passivoElim += e;
    }

    function emitExtras(rootKey) {
      emitTrees(trees.map(t => t.extras[rootKey]), 0);
    }

    pushHeader('ATIVO', 'root');
    emitSection('AC'); emitSection('ANC');
    emitExtras('A');
    push('TOTAL DO ATIVO', 'grandtotal', trees.map(t => t.ativo), ativoElim);
    pushHeader('PASSIVO E PATRIMÔNIO LÍQUIDO', 'root');
    emitSection('PC'); emitSection('PNC'); emitSection('PL');
    emitExtras('P');
    push('TOTAL DO PASSIVO + PL', 'grandtotal', trees.map(t => t.passivo), passivoElim);

    const totalAtivo = trees.reduce((s, t) => s + t.ativo, 0) + ativoElim;
    const totalPassivo = trees.reduce((s, t) => s + t.passivo, 0) + passivoElim;
    const diff = totalAtivo - totalPassivo;
    if (Math.abs(residual) > 0.02) {
      push('SALDO INTERCOMPANY RESIDUAL NÃO CONCILIADO (ver nota)', 'alerta', new Array(n).fill(0), 0);
      out[out.length - 1].consolidado = residual;
    }

    return { rows: out, totalAtivo, totalPassivo, diff, fecha: Math.abs(diff) <= 0.02, residual };
  }

  // ---------- Consolidação da DRE (template canônico fixo) ----------
  // A DRE NÃO usa merge posicional de árvores: a hierarquia segue um template
  // canônico fixo (Receita Bruta → Deduções → Receita Líquida → Custos →
  // ROL → Despesas Operacionais → Resultado Financeiro → ROL → Receitas Não
  // Operacionais → Lucro Líquido). Blocos/contas das duas empresas são casados
  // pela descrição normalizada (U.dreKey); analíticas de A primeiro (ordem
  // original), depois as que só existem em B. Subtotais saem por fórmula
  // (campo `comps` = índices das linhas componentes).
  function num(v) { return (typeof v === 'number' && isFinite(v)) ? v : 0; }

  const DRE_COMPUTED = /^(RECEITA LIQUIDA|LUCRO BRUTO|RESULTADO OPERACIONAL|RESULTADO ANTES|RESULTADO LIQUIDO|LUCRO LIQUIDO|PREJUIZO|RESULTADO DO EXERC)/;

  // Classifica os blocos de dados de uma empresa nas seções canônicas.
  function dreStruct(dre) {
    const data = (dre.blocks || []).filter(b => !DRE_COMPUTED.test(b.key));
    const m = { receitaBruta: null, deducoes: null, custos: [], despesas: [], financeiro: [], naoOper: [], extras: [] };
    let i = 0;
    while (i < data.length) {
      const b = data[i], k = b.key;
      if (/^RECEITAS? BRUTA/.test(k)) { m.receitaBruta = b; i++; continue; }
      if (/^DEDUC/.test(k)) { m.deducoes = b; i++; continue; }
      if (/^CUSTOS$/.test(k) || /^DESPESAS OPERACIONAIS$/.test(k)) {
        // Grupo contêiner: consome os sub-blocos seguintes até a soma fechar
        // com o valor do grupo (membros podem variar por empresa).
        const list = /^CUSTOS$/.test(k) ? m.custos : m.despesas;
        let sum = 0; i++;
        while (i < data.length && Math.abs(sum - b.value) > 0.02) { sum += data[i].value; list.push(data[i]); i++; }
        continue;
      }
      if (/FINANCEIR/.test(k)) { m.financeiro.push(b); i++; continue; }
      if (/NAO OPERACION/.test(k) && /RECEITA/.test(k)) { m.naoOper.push(b); i++; continue; }
      if (/^CUSTOS /.test(k) || /^CMV/.test(k)) { m.custos.push(b); i++; continue; }     // custo fora do grupo pai
      if (/^DESPESAS /.test(k)) { m.despesas.push(b); i++; continue; }                    // despesa fora do grupo pai
      m.extras.push(b); i++;                                                              // demais blocos operacionais
    }
    return m;
  }

  // Une as listas de blocos das empresas por chave: ordem canônica preferida
  // (regex, com sinônimos de grafia) → demais chaves na ordem em que aparecem,
  // empresa por empresa.
  function mergeBlockLists(lists, prefOrder) {
    const map = new Map();
    const put = (b, ci) => {
      if (!map.has(b.key)) map.set(b.key, { key: b.key, descricao: b.descricao, by: new Array(lists.length).fill(null) });
      map.get(b.key).by[ci] = b;
    };
    lists.forEach((list, ci) => list.forEach(b => put(b, ci)));
    const keys = [];
    lists.forEach(list => list.forEach(b => { if (keys.indexOf(b.key) < 0) keys.push(b.key); }));
    const ordered = [], seen = new Set();
    (prefOrder || []).forEach(re => keys.forEach(k => {
      if (!seen.has(k) && re.test(k)) { ordered.push(map.get(k)); seen.add(k); }
    }));
    keys.forEach(k => { if (!seen.has(k)) { ordered.push(map.get(k)); seen.add(k); } });
    return ordered;
  }

  // Une as analíticas de um bloco casado: a 1ª empresa dita a ordem, as
  // seguintes somam na linha de mesma chave ou abrem linha nova no fim.
  function mergeAnalytics(merged, n) {
    const out = [], idx = new Map();
    merged.by.forEach((blk, ci) => {
      (blk ? blk.analytics : []).forEach(a => {
        if (ci > 0 && idx.has(a.key)) { out[idx.get(a.key)].vals[ci] += num(a.value); return; }
        if (ci === 0 || !idx.has(a.key)) idx.set(a.key, out.length);
        const vals = new Array(n).fill(0);
        vals[ci] = num(a.value);
        out.push({ descricao: a.descricao, vals: vals });
      });
    });
    return out;
  }

  const ORDER_CUSTOS = [/^CUSTOS GERAIS/, /^CUSTOS COM PESSOAL/];
  const ORDER_DESP = [/^DESPESAS (COM MARKETING|COMERCIA)/, /^DESPESAS ADMINISTRA/,
    /^DESPESAS COM TECNOLOG/, /^DESPESAS COM INFRAESTRUT/];

  function consolidateDRE(emps) {
    checkEmps(emps);
    const n = emps.length;
    const S = emps.map(e => dreStruct(e.dre));
    // Primeiro valor não-nulo entre as empresas (a 1ª que tiver o bloco).
    const firstOf = sel => { for (let i = 0; i < S.length; i++) { const v = sel(S[i]); if (v) return v; } return null; };
    const lines = [];
    const row = (desc, kind, vals, comps) => {
      const v = new Array(n);
      for (let i = 0; i < n; i++) v[i] = num(vals[i]);
      const soma = v.reduce((s, x) => s + x, 0);
      lines.push({ desc, kind, vals: v, elim: 0, soma, consolidado: soma, comps: comps || null });
      return lines.length - 1;
    };
    // Bloco sintético: linha do bloco (fórmula = soma das analíticas) + analíticas.
    const emitBlock = (merged, kind) => {
      const an = mergeAnalytics(merged, n);
      const vals = new Array(n);
      for (let ci = 0; ci < n; ci++) {
        vals[ci] = an.length ? an.reduce((s, a) => s + a.vals[ci], 0)
          : (merged.by[ci] ? num(merged.by[ci].value) : 0);
      }
      const idx = row(merged.descricao, kind || 'subtotal', vals, null);
      const comps = an.map(a => row('   ' + a.descricao, 'item', a.vals));
      if (comps.length) lines[idx].comps = comps;
      return idx;
    };
    const sumRows = idxs => {
      const acc = new Array(n).fill(0);
      idxs.forEach(i => { for (let c = 0; c < n; c++) acc[c] += lines[i].vals[c]; });
      return acc;
    };

    // 1–2) Receita Bruta e Deduções
    const rbIdx = emitBlock({ descricao: (firstOf(s => s.receitaBruta) || { descricao: 'RECEITA BRUTA' }).descricao,
      by: S.map(s => s.receitaBruta) });
    const dedIdx = emitBlock({ descricao: (firstOf(s => s.deducoes) || { descricao: 'DEDUÇÕES' }).descricao,
      by: S.map(s => s.deducoes) });
    // 3) Receita Líquida (fórmula)
    const rlIdx = row('(=) RECEITA LÍQUIDA', 'total', sumRows([rbIdx, dedIdx]), [rbIdx, dedIdx]);
    // 4) Custos (subgrupos na ordem canônica, depois subtotal por fórmula)
    const custosM = mergeBlockLists(S.map(s => s.custos), ORDER_CUSTOS);
    const cIdxs = custosM.map(mb => emitBlock(mb));
    let custosIdx = -1;
    if (cIdxs.length) custosIdx = row('CUSTOS', 'subtotal', sumRows(cIdxs), cIdxs);
    // 5) ROL após custos (fórmula)
    const rol5comps = custosIdx >= 0 ? [rlIdx, custosIdx] : [rlIdx];
    const rol5Idx = row('(=) RESULTADO OPERACIONAL LÍQUIDO', 'total', sumRows(rol5comps), rol5comps);
    // 6) Despesas operacionais
    const despM = mergeBlockLists(S.map(s => s.despesas), ORDER_DESP);
    const dIdxs = despM.map(mb => emitBlock(mb));
    let despIdx = -1;
    if (dIdxs.length) despIdx = row('DESPESAS OPERACIONAIS', 'subtotal', sumRows(dIdxs), dIdxs);
    // 7) Resultado financeiro (+ blocos operacionais extras, fiéis à origem)
    const finIdxs = mergeBlockLists(S.map(s => s.financeiro), []).map(mb => emitBlock(mb));
    const extraIdxs = mergeBlockLists(S.map(s => s.extras), []).map(mb => emitBlock(mb));
    // 8) ROL após despesas (fórmula)
    const rol8comps = [rol5Idx].concat(despIdx >= 0 ? [despIdx] : []).concat(finIdxs).concat(extraIdxs);
    const rol8Idx = rol8comps.length > 1
      ? row('(=) RESULTADO OPERACIONAL LÍQUIDO', 'total', sumRows(rol8comps), rol8comps) : rol5Idx;
    // 9) Receitas não operacionais (somente se existir em alguma empresa)
    const naoIdxs = mergeBlockLists(S.map(s => s.naoOper), []).map(mb => emitBlock(mb));
    // 10) Lucro líquido (fórmula)
    const llComps = [rol8Idx].concat(naoIdxs);
    const llIdx = row('(=) LUCRO LÍQUIDO DO EXERCÍCIO', 'grandtotal', sumRows(llComps), llComps);

    // ----- Validação obrigatória (tolerância ≤ R$ 0,02) -----
    const checks = [];
    const chk = (empresa, conta, calc, oficial) => {
      if (oficial === undefined || oficial === null) return;
      const dif = calc - oficial;
      checks.push({ empresa, conta, calc, oficial, dif, ok: Math.abs(dif) <= 0.02 });
    };
    emps.forEach((emp, ci) => {
      const o = emp.dre.official || {};
      chk(emp.id, 'Receita Líquida', lines[rlIdx].vals[ci], o.receitaLiquida);
      chk(emp.id, 'Resultado Operacional Líquido (após custos)', lines[rol5Idx].vals[ci], o.lucroBruto);
      chk(emp.id, 'Resultado Operacional Líquido (após despesas)', lines[rol8Idx].vals[ci], o.resultadoOperacional);
      chk(emp.id, 'Lucro Líquido do Exercício', lines[llIdx].vals[ci], o.resultLiquido);
    });
    // O resultado exibido no consolidado é sempre o recalculado pelas linhas
    // da DRE. Somar "oficiais" ausentes como zero mascara resultado e gera a
    // falsa divergência exibida na tela.
    const resLiq = emps.map((e, ci) => lines[llIdx].vals[ci]);
    const resultLiquido = lines[llIdx].consolidado;
    const oficiais = emps.map(e => (e.dre.official || {}).resultLiquido);
    const temOficial = oficiais.map(v => typeof v === 'number' && isFinite(v));
    const todosOficiais = temOficial.every(Boolean);
    const resLiqEsperado = todosOficiais ? oficiais.reduce((s, v) => s + v, 0) : undefined;
    if (resLiqEsperado !== undefined)
      chk('Consolidado', 'Lucro Líquido = ' + emps.map(e => e.id).join(' + '), resultLiquido, resLiqEsperado);
    const warnings = [];
    if (!todosOficiais) {
      const faltam = emps.filter((e, ci) => !temOficial[ci]).map(e => 'Empresa ' + e.id);
      warnings.push('O Resultado Líquido oficial não foi identificado em ' +
        (faltam.length === emps.length ? (emps.length === 2 ? 'ambas as DREs' : 'todas as DREs') : faltam.join(', ')) +
        '; o consolidado exibe o valor recalculado pelas contas analíticas.');
    }
    const errors = checks.filter(c => !c.ok)
      .map(c => `DRE ${c.empresa}: ${c.conta} recalculado (${U.fmt(c.calc)}) difere do oficial (${U.fmt(c.oficial)}) em ${U.fmt(c.dif)}`);

    // ----- Auditoria: sintética == soma das analíticas (por empresa) -----
    const blockAudit = [];
    emps.forEach(emp => {
      (emp.dre.blocks || []).forEach(bk => {
        if (!bk.analytics.length) return;
        const sum = bk.analytics.reduce((s, x) => s + num(x.value), 0);
        blockAudit.push({ empresa: emp.id, conta: bk.descricao, sintetico: bk.value, soma: sum,
          dif: bk.value - sum, ok: Math.abs(bk.value - sum) <= 0.02 });
      });
    });

    const bate = errors.length === 0;
    return { lines, resLiq, resLiqEsperado, resultLiquido, bate, errors, warnings, checks, blockAudit };
  }

  // ----- Auditoria do balanço: cada sintética == soma dos filhos diretos -----
  function auditBalance(emp) {
    const rows = emp.balanco.rows || [], out = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r.synthetic) continue;
      let sum = 0, has = false;
      for (let j = i + 1; j < rows.length; j++) {
        if (rows[j].level <= r.level) break;
        if (rows[j].level === r.level + 1) { sum += rows[j].saldoAtual; has = true; }
      }
      if (!has) continue;
      const dif = r.saldoAtual - sum;
      out.push({ conta: r.descricao, sintetico: r.saldoAtual, soma: sum, dif, ok: Math.abs(dif) <= 0.02 });
    }
    return out;
  }

  global.C = { detectIntercompany, consolidateBalance, consolidateDRE, financialPositions,
    auditBalance, sideOf, MAX_EMPRESAS, MIN_EMPRESAS };
})(typeof window !== 'undefined' ? window : globalThis);
