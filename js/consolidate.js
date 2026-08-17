/* consolidate.js — intercompany (detecção/divergência), consolidação do balanço
   por seções (AC/ANC/PC/PNC/PL) e DRE gerencial (§3.2). */
(function (global) {
  'use strict';
  const U = global.U;

  const NAME_THRESHOLD = 0.6;

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

  function financialPositions(empA, empB, pairs) {
    const eliminations = positionEliminations(pairs);
    const positions = {};
    Object.keys(POSITION_TYPES).forEach(key => {
      const type = POSITION_TYPES[key], byDesc = new Map();
      const add = (item, company) => {
        const canon = U.norm(item.descricao);
        if (!byDesc.has(canon)) byDesc.set(canon, { desc: item.descricao, valA: 0, valB: 0, elim: 0 });
        const line = byDesc.get(canon);
        line[company] += item.value;
        line.elim += eliminations.get(item.rowId) || 0;
      };
      positionRows(empA, type).forEach(item => add(item, 'valA'));
      positionRows(empB, type).forEach(item => add(item, 'valB'));
      const lines = Array.from(byDesc.values()).map(line => {
        line.soma = line.valA + line.valB;
        line.consolidado = line.soma + line.elim;
        return line;
      }).sort((a, b) => a.desc.localeCompare(b.desc, 'pt-BR'));
      positions[key] = {
        title: type.title, lines,
        totalA: lines.reduce((s, x) => s + x.valA, 0),
        totalB: lines.reduce((s, x) => s + x.valB, 0),
        totalElim: lines.reduce((s, x) => s + x.elim, 0),
        total: lines.reduce((s, x) => s + x.consolidado, 0)
      };
    });
    return positions;
  }

  // ---------- Detecção intercompany ----------
  function detectIntercompany(empA, empB) {
    const aAll = analytics(empA), bAll = analytics(empB);
    const nameA = empA.balanco.header.nome, nameB = empB.balanco.header.nome;
    const cnpjA = U.digits(empA.balanco.header.cnpj), cnpjB = U.digits(empB.balanco.header.cnpj);

    const candA = aAll.filter(r => U.nameAppears(r.descricao, nameB) >= NAME_THRESHOLD ||
      (cnpjB && U.digits(r.descricao).indexOf(cnpjB) >= 0))
      .map(r => ({ company: 'A', rowId: r.rowId, desc: r.descricao, value: r.saldoAtual, side: sideOf(r.descricao) }));
    const candB = bAll.filter(r => U.nameAppears(r.descricao, nameA) >= NAME_THRESHOLD ||
      (cnpjA && U.digits(r.descricao).indexOf(cnpjA) >= 0))
      .map(r => ({ company: 'B', rowId: r.rowId, desc: r.descricao, value: r.saldoAtual, side: sideOf(r.descricao) }));

    const pairs = []; const usedA = new Set(), usedB = new Set();
    candA.forEach((a, ia) => {
      let best = -1, bestScore = -1;
      candB.forEach((b, ib) => {
        if (usedB.has(ib)) return;
        let sc = 1;
        if ((a.side === 'receber' && b.side === 'pagar') || (a.side === 'pagar' && b.side === 'receber')) sc += 2;
        if (sc > bestScore) { bestScore = sc; best = ib; }
      });
      if (best >= 0) {
        usedA.add(ia); usedB.add(best);
        const b = candB[best];
        pairs.push({ a, b, valA: a.value, valB: b.value,
          confirmavel: Math.min(Math.abs(a.value), Math.abs(b.value)),
          divergencia: Math.abs(Math.abs(a.value) - Math.abs(b.value)),
          decisao: 'eliminar', tratamento: 'residual' });
      }
    });
    return { pairs, soloA: candA.filter((a, i) => !usedA.has(i)), soloB: candB.filter((b, i) => !usedB.has(i)) };
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

  // Força o desdobramento de grupos genéricos (ex.: "Disponível" e
  // "Obrigações") até a conta patrimonial que o usuário precisa enxergar.
  // Assim Bancos, Clientes/Contas a Receber e Fornecedores aparecem dentro
  // do próprio balancete, sem criar demonstrativos paralelos ou duplicar saldo.
  const BALANCETE_HIGHLIGHTS = /\b(CLIENTES?|DUPLICATAS? A RECEBER|CONTAS? A RECEBER|TITULOS? A RECEBER|FORNECEDORES?|DUPLICATAS? A PAGAR|BANCOS?|CONTA CORRENTE)\b/;

  function highlightKind(norm) {
    if (/\b(CLIENTES?|DUPLICATAS? A RECEBER|CONTAS? A RECEBER|TITULOS? A RECEBER)\b/.test(norm)) return 'clientes';
    if (/\b(FORNECEDORES?|DUPLICATAS? A PAGAR)\b/.test(norm)) return 'fornecedores';
    if (/\b(BANCOS?|CONTA CORRENTE)\b/.test(norm)) return 'bancos';
    return null;
  }

  // Emite os itens de uma seção, drilando contêineres que tenham intercompany.
  function sectionItems(rows, l1Idx, sign, other, empId, secKey) {
    const items = [];
    const l1Level = rows[l1Idx].level;
    function processNode(start, end, node, inheritedHighlight, inheritedGroupLabel) {
      // intercompany no subtree?
      let icIdx = -1;
      for (let j = start; j < end; j++) {
        if (!rows[j].synthetic && U.nameAppears(rows[j].descricao, other) >= NAME_THRESHOLD) { icIdx = j; break; }
      }
      if (icIdx < 0 && !node.synthetic && U.nameAppears(node.descricao, other) >= NAME_THRESHOLD) icIdx = start;
      let hasHighlightBelow = false;
      for (let j = start; j < end; j++) {
        if (BALANCETE_HIGHLIGHTS.test(U.norm(rows[j].descricao))) { hasHighlightBelow = true; break; }
      }
      const nodeHighlight = highlightKind(U.norm(node.descricao));
      const activeHighlight = inheritedHighlight || nodeHighlight || null;
      // Só uma conta sintética abre um grupo. Uma analítica como "Cliente
      // Alfa" pode conter a palavra Cliente, mas deve ficar abaixo do grupo
      // sintético "Clientes", e não criar um grupo novo.
      const activeGroupLabel = inheritedGroupLabel ||
        (nodeHighlight && node.synthetic ? node.descricao : null);
      const nodeIsHighlight = !!nodeHighlight;
      const mixedIntercompany = icIdx >= 0 && Math.abs(node.saldoAtual - rows[icIdx].saldoAtual) > 0.02;
      // Contêiner misto com intercompany, ou agrupador genérico que esconde uma
      // conta destacada (Disponível → Bancos, Obrigações → Fornecedores): desce
      // um nível, mantendo o saldo apenas uma vez no balanço.
      // Nas contas destacadas, continua até as analíticas: "Clientes" deixa
      // de ser uma linha única e passa a listar cada cliente efetivamente em
      // aberto; a mesma regra vale para fornecedores e bancos.
      const needsAnalyticDetail = !!activeHighlight && node.synthetic;
      if ((mixedIntercompany || (hasHighlightBelow && !nodeIsHighlight) || needsAnalyticDetail) && (end - start) > 1) {
        let k = start + 1;
        while (k < end && rows[k].level > node.level) {
          if (rows[k].level === node.level + 1) {
            let cend = k + 1; while (cend < end && rows[cend].level > rows[k].level) cend++;
            processNode(k, cend, rows[k], activeHighlight, activeGroupLabel); k = cend;
          } else k++;
        }
      } else {
        const isInter = icIdx >= 0;
        items.push({
          label: isInter ? rows[icIdx].descricao : node.descricao,
          canon: isInter ? (secKey === 'AC' || secKey === 'ANC' ? 'IC_RECEBER' : 'IC_PAGAR') : U.norm(node.descricao),
          value: sign * node.saldoAtual, isInter,
          highlight: activeHighlight,
          highlightGroup: activeGroupLabel ? {
            key: activeHighlight + '|' + U.norm(activeGroupLabel), label: activeGroupLabel
          } : null,
          rowId: empId + ':' + (isInter ? rows[icIdx].rowIndex : node.rowIndex)
        });
      }
    }
    let i = l1Idx + 1;
    while (i < rows.length && rows[i].level > l1Level) {
      if (rows[i].level === l1Level + 1) {
        let end = i + 1; while (end < rows.length && rows[end].level > rows[i].level) end++;
        processNode(i, end, rows[i]); i = end;
      } else i++;
    }
    return items;
  }

  function buildSections(emp, other) {
    const rows = emp.balanco.rows;
    const model = {}; SEC_ORDER.forEach(s => model[s] = { total: 0, items: [] });
    // Grupos nível-1 fora das seções conhecidas (ex.: bancos/clientes/fornecedores
    // em aberto, conta de resultado ainda aberta no balanço): o valor já está no
    // total-raiz do arquivo, então não podem sumir do relatório.
    const extras = [];
    let root = null;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i], n = U.norm(r.descricao);
      if (r.level === 0) { root = /PASSIVO/.test(n) ? 'P' : 'A'; continue; }
      if (r.level === 1) {
        const sec = classL1(n, root);
        if (sec) {
          const sign = (sec === 'PC' || sec === 'PNC' || sec === 'PL') ? -1 : 1;
          model[sec].total += sign * r.saldoAtual;
          model[sec].items = model[sec].items.concat(sectionItems(rows, i, sign, other, emp.id, sec));
        } else if (r.saldoAtual !== 0 && !/^TOTAL/.test(n)) {
          const sign = root === 'P' ? -1 : 1;
          extras.push({ label: r.descricao, canon: n, value: sign * r.saldoAtual,
            isInter: false, root: root, rowId: emp.id + ':' + r.rowIndex });
        }
      }
    }
    const ativo = (emp.balanco.totalAtivo || 0);
    const passivo = -(emp.balanco.totalPassivo || 0);
    return { model, ativo, passivo, extras };
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

  function consolidateBalance(empA, empB, pairs) {
    const A = balanceTree(empA), B = balanceTree(empB);

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
    const push = (desc, kind, valA, valB, e) => {
      const soma = (valA || 0) + (valB || 0); const el = e || 0;
      out.push({ desc, kind, valA: valA || 0, valB: valB || 0, elim: el, soma, consolidado: soma + el });
    };
    const pushHeader = (desc, kind) => out.push({ desc, kind, header: true });

    function eliminationOf(node) {
      if (!node) return 0;
      return (rowElim.get(node.rowId) || 0) +
        (node.children || []).reduce((s, child) => s + eliminationOf(child), 0);
    }

    function emitTrees(aNodes, bNodes, depth) {
      const bBuckets = new Map();
      bNodes.forEach((node, i) => {
        if (!bBuckets.has(node.key)) bBuckets.set(node.key, []);
        bBuckets.get(node.key).push({ node, i });
      });
      const usedB = new Set();
      const emitPair = (aNode, bNode) => {
        const childA = aNode ? aNode.children || [] : [];
        const childB = bNode ? bNode.children || [] : [];
        const e = eliminationOf(aNode) + eliminationOf(bNode);
        const hasChildren = childA.length || childB.length;
        const label = (aNode || bNode).label;
        push('   '.repeat(depth + 1) + label,
          hasChildren ? 'subtotal' : (Math.abs(e) > 0.000001 ? 'inter' : 'item'),
          aNode ? aNode.value : 0, bNode ? bNode.value : 0, e);
        if (hasChildren) emitTrees(childA, childB, depth + 1);
      };
      aNodes.forEach(aNode => {
        const bucket = bBuckets.get(aNode.key) || [];
        const match = bucket.find(x => !usedB.has(x.i));
        if (match) usedB.add(match.i);
        emitPair(aNode, match ? match.node : null);
      });
      // Contas exclusivas da Empresa B entram dentro do mesmo pai e na ordem
      // original de B, antes de o relatório avançar para a próxima sintética.
      bNodes.forEach((bNode, i) => { if (!usedB.has(i)) emitPair(null, bNode); });
    }

    let ativoElim = 0, passivoElim = 0;
    function emitSection(secKey) {
      const a = A.sections[secKey], b = B.sections[secKey];
      pushHeader(SEC_NAME[secKey], 'section');
      emitTrees(a.nodes, b.nodes, 0);
      const e = a.nodes.reduce((s, x) => s + eliminationOf(x), 0) +
        b.nodes.reduce((s, x) => s + eliminationOf(x), 0);
      push('Total do ' + titleCase(SEC_NAME[secKey]), 'subtotal', a.total, b.total, e);
      if (secKey === 'AC' || secKey === 'ANC') ativoElim += e; else passivoElim += e;
    }

    function emitExtras(rootKey) {
      emitTrees(A.extras[rootKey], B.extras[rootKey], 0);
    }

    pushHeader('ATIVO', 'root');
    emitSection('AC'); emitSection('ANC');
    emitExtras('A');
    push('TOTAL DO ATIVO', 'grandtotal', A.ativo, B.ativo, ativoElim);
    pushHeader('PASSIVO E PATRIMÔNIO LÍQUIDO', 'root');
    emitSection('PC'); emitSection('PNC'); emitSection('PL');
    emitExtras('P');
    push('TOTAL DO PASSIVO + PL', 'grandtotal', A.passivo, B.passivo, passivoElim);

    const totalAtivo = A.ativo + B.ativo + ativoElim;
    const totalPassivo = A.passivo + B.passivo + passivoElim;
    const diff = totalAtivo - totalPassivo;
    if (Math.abs(residual) > 0.02) {
      push('SALDO INTERCOMPANY RESIDUAL NÃO CONCILIADO (ver nota)', 'alerta', 0, 0, 0);
      out[out.length - 1].consolidado = residual;
    }

    return { rows: out, totalAtivo, totalPassivo, diff, fecha: Math.abs(diff) <= 0.02, residual };
  }

  function labelFor(it, nameA, nameB) {
    if (it.canon === 'IC_RECEBER') return 'Mútuos/valores a receber — ' + firstWord(nameB) + ' (intercompany)';
    if (it.canon === 'IC_PAGAR') return 'Empréstimo/valores a pagar — ' + firstWord(nameB) + ' (intercompany)';
    if (it.highlightGroup) return '      ' + titleCase(it.label);
    if (it.highlight) {
      const prefix = { clientes: 'Clientes em aberto', fornecedores: 'Fornecedores em aberto', bancos: 'Bancos' }[it.highlight];
      return '      ' + prefix + ' — ' + titleCase(it.label);
    }
    const n = U.norm(it.label);
    if (/\b(CLIENTES?|DUPLICATAS? A RECEBER|CONTAS? A RECEBER|TITULOS? A RECEBER)\b/.test(n)) return '   Clientes em aberto';
    if (/\b(FORNECEDORES?|DUPLICATAS? A PAGAR)\b/.test(n)) return '   Fornecedores em aberto';
    if (/\b(BANCOS?|CONTA CORRENTE)\b/.test(n)) return '   Bancos';
    return '   ' + titleCase(it.label);
  }
  function firstWord(s) { return String(s || '').trim().split(/\s+/)[0] || 'coligada'; }

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

  // Une listas de blocos A/B por chave: ordem canônica preferida (regex, com
  // sinônimos de grafia) → demais de A na ordem original → só-B na ordem de B.
  function mergeBlockLists(listA, listB, prefOrder) {
    const map = new Map();
    const put = (b, side) => {
      if (!map.has(b.key)) map.set(b.key, { key: b.key, descricao: b.descricao, A: null, B: null });
      map.get(b.key)[side] = b;
    };
    listA.forEach(b => put(b, 'A')); listB.forEach(b => put(b, 'B'));
    const keys = [];
    listA.forEach(b => { if (keys.indexOf(b.key) < 0) keys.push(b.key); });
    listB.forEach(b => { if (keys.indexOf(b.key) < 0) keys.push(b.key); });
    const ordered = [], seen = new Set();
    (prefOrder || []).forEach(re => keys.forEach(k => {
      if (!seen.has(k) && re.test(k)) { ordered.push(map.get(k)); seen.add(k); }
    }));
    keys.forEach(k => { if (!seen.has(k)) { ordered.push(map.get(k)); seen.add(k); } });
    return ordered;
  }

  // Une as analíticas de um bloco casado (A primeiro, depois só-B).
  function mergeAnalytics(merged) {
    const aList = merged.A ? merged.A.analytics : [], bList = merged.B ? merged.B.analytics : [];
    const out = [], idx = new Map();
    aList.forEach(a => { idx.set(a.key, out.length); out.push({ descricao: a.descricao, valA: num(a.value), valB: 0 }); });
    bList.forEach(b => {
      if (idx.has(b.key)) out[idx.get(b.key)].valB += num(b.value);
      else out.push({ descricao: b.descricao, valA: 0, valB: num(b.value) });
    });
    return out;
  }

  const ORDER_CUSTOS = [/^CUSTOS GERAIS/, /^CUSTOS COM PESSOAL/];
  const ORDER_DESP = [/^DESPESAS (COM MARKETING|COMERCIA)/, /^DESPESAS ADMINISTRA/,
    /^DESPESAS COM TECNOLOG/, /^DESPESAS COM INFRAESTRUT/];

  function consolidateDRE(empA, empB) {
    const SA = dreStruct(empA.dre), SB = dreStruct(empB.dre);
    const lines = [];
    const row = (desc, kind, valA, valB, comps) => {
      const va = num(valA), vb = num(valB);
      lines.push({ desc, kind, valA: va, valB: vb, elim: 0, soma: va + vb, consolidado: va + vb, comps: comps || null });
      return lines.length - 1;
    };
    // Bloco sintético: linha do bloco (fórmula = soma das analíticas) + analíticas.
    const emitBlock = (merged, kind) => {
      const an = mergeAnalytics(merged);
      const sumA = an.reduce((s, a) => s + a.valA, 0), sumB = an.reduce((s, a) => s + a.valB, 0);
      const vA = merged.A ? (an.length ? sumA : num(merged.A.value)) : (an.length ? sumA : 0);
      const vB = merged.B ? (an.length ? sumB : num(merged.B.value)) : (an.length ? sumB : 0);
      const idx = row(merged.descricao, kind || 'subtotal', vA, vB, null);
      const comps = an.map(a => row('   ' + a.descricao, 'item', a.valA, a.valB));
      if (comps.length) lines[idx].comps = comps;
      return idx;
    };
    const sumRows = idxs => idxs.reduce((s, i) => [s[0] + lines[i].valA, s[1] + lines[i].valB], [0, 0]);

    // 1–2) Receita Bruta e Deduções
    const rbIdx = emitBlock({ descricao: (SA.receitaBruta || SB.receitaBruta || { descricao: 'RECEITA BRUTA' }).descricao,
      A: SA.receitaBruta, B: SB.receitaBruta });
    const dedIdx = emitBlock({ descricao: (SA.deducoes || SB.deducoes || { descricao: 'DEDUÇÕES' }).descricao,
      A: SA.deducoes, B: SB.deducoes });
    // 3) Receita Líquida (fórmula)
    let [a, b] = sumRows([rbIdx, dedIdx]);
    const rlIdx = row('(=) RECEITA LÍQUIDA', 'total', a, b, [rbIdx, dedIdx]);
    // 4) Custos (subgrupos na ordem canônica, depois subtotal por fórmula)
    const custosM = mergeBlockLists(SA.custos, SB.custos, ORDER_CUSTOS);
    const cIdxs = custosM.map(mb => emitBlock(mb));
    let custosIdx = -1;
    if (cIdxs.length) { [a, b] = sumRows(cIdxs); custosIdx = row('CUSTOS', 'subtotal', a, b, cIdxs); }
    // 5) ROL após custos (fórmula)
    const rol5comps = custosIdx >= 0 ? [rlIdx, custosIdx] : [rlIdx];
    [a, b] = sumRows(rol5comps);
    const rol5Idx = row('(=) RESULTADO OPERACIONAL LÍQUIDO', 'total', a, b, rol5comps);
    // 6) Despesas operacionais
    const despM = mergeBlockLists(SA.despesas, SB.despesas, ORDER_DESP);
    const dIdxs = despM.map(mb => emitBlock(mb));
    let despIdx = -1;
    if (dIdxs.length) { [a, b] = sumRows(dIdxs); despIdx = row('DESPESAS OPERACIONAIS', 'subtotal', a, b, dIdxs); }
    // 7) Resultado financeiro (+ blocos operacionais extras, fiéis à origem)
    const finIdxs = mergeBlockLists(SA.financeiro, SB.financeiro, []).map(mb => emitBlock(mb));
    const extraIdxs = mergeBlockLists(SA.extras, SB.extras, []).map(mb => emitBlock(mb));
    // 8) ROL após despesas (fórmula)
    const rol8comps = [rol5Idx].concat(despIdx >= 0 ? [despIdx] : []).concat(finIdxs).concat(extraIdxs);
    [a, b] = sumRows(rol8comps);
    const rol8Idx = rol8comps.length > 1 ? row('(=) RESULTADO OPERACIONAL LÍQUIDO', 'total', a, b, rol8comps) : rol5Idx;
    // 9) Receitas não operacionais (somente se existir em A ou B)
    const naoIdxs = mergeBlockLists(SA.naoOper, SB.naoOper, []).map(mb => emitBlock(mb));
    // 10) Lucro líquido (fórmula)
    const llComps = [rol8Idx].concat(naoIdxs);
    [a, b] = sumRows(llComps);
    const llIdx = row('(=) LUCRO LÍQUIDO DO EXERCÍCIO', 'grandtotal', a, b, llComps);

    // ----- Validação obrigatória (tolerância ≤ R$ 0,02) -----
    const checks = [];
    const chk = (empresa, conta, calc, oficial) => {
      if (oficial === undefined || oficial === null) return;
      const dif = calc - oficial;
      checks.push({ empresa, conta, calc, oficial, dif, ok: Math.abs(dif) <= 0.02 });
    };
    [['A', empA, 'valA'], ['B', empB, 'valB']].forEach(([rot, emp, col]) => {
      const o = emp.dre.official || {};
      chk(rot, 'Receita Líquida', lines[rlIdx][col], o.receitaLiquida);
      chk(rot, 'Resultado Operacional Líquido (após custos)', lines[rol5Idx][col], o.lucroBruto);
      chk(rot, 'Resultado Operacional Líquido (após despesas)', lines[rol8Idx][col], o.resultadoOperacional);
      chk(rot, 'Lucro Líquido do Exercício', lines[llIdx][col], o.resultLiquido);
    });
    // O resultado exibido no consolidado é sempre o recalculado pelas linhas
    // da DRE. Somar "oficiais" ausentes como zero mascara resultado e gera a
    // falsa divergência exibida na tela.
    const resLiqA = lines[llIdx].valA, resLiqB = lines[llIdx].valB;
    const resultLiquido = lines[llIdx].consolidado;
    const oficialA = (empA.dre.official || {}).resultLiquido;
    const oficialB = (empB.dre.official || {}).resultLiquido;
    const temOficialA = typeof oficialA === 'number' && isFinite(oficialA);
    const temOficialB = typeof oficialB === 'number' && isFinite(oficialB);
    const resLiqEsperado = temOficialA && temOficialB ? oficialA + oficialB : undefined;
    if (resLiqEsperado !== undefined)
      chk('Consolidado', 'Lucro Líquido = A + B', resultLiquido, resLiqEsperado);
    const warnings = [];
    if (!temOficialA || !temOficialB)
      warnings.push('O Resultado Líquido oficial não foi identificado em ' +
        (!temOficialA && !temOficialB ? 'ambas as DREs' : (!temOficialA ? 'Empresa A' : 'Empresa B')) +
        '; o consolidado exibe o valor recalculado pelas contas analíticas.');
    const errors = checks.filter(c => !c.ok)
      .map(c => `DRE ${c.empresa}: ${c.conta} recalculado (${U.fmt(c.calc)}) difere do oficial (${U.fmt(c.oficial)}) em ${U.fmt(c.dif)}`);

    // ----- Auditoria: sintética == soma das analíticas (por empresa) -----
    const blockAudit = [];
    [['A', empA], ['B', empB]].forEach(([rot, emp]) => {
      (emp.dre.blocks || []).forEach(bk => {
        if (!bk.analytics.length) return;
        const sum = bk.analytics.reduce((s, x) => s + num(x.value), 0);
        blockAudit.push({ empresa: rot, conta: bk.descricao, sintetico: bk.value, soma: sum,
          dif: bk.value - sum, ok: Math.abs(bk.value - sum) <= 0.02 });
      });
    });

    const bate = errors.length === 0;
    return { lines, resLiqA, resLiqB, resLiqEsperado, resultLiquido, bate, errors, warnings, checks, blockAudit };
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

  global.C = { detectIntercompany, consolidateBalance, consolidateDRE, financialPositions, auditBalance, sideOf };
})(typeof window !== 'undefined' ? window : globalThis);
