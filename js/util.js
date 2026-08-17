/* util.js — funções de base: números pt-BR, normalização, similaridade. */
(function (global) {
  'use strict';

  // Converte célula (number ou texto pt-BR) em número JS. Mantém o sinal.
  // Aceita: 1234.56 | "1.234,56" | "(1.234,56)" | "-1.234,56" | "1234,56" | "R$ 1.234,56"
  function parseNum(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    let s = String(v).trim();
    if (!s) return 0;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    if (/^-/.test(s.trim())) neg = true;
    // Mantém apenas dígitos, vírgula e ponto.
    s = s.replace(/[^0-9.,]/g, '');
    if (!s) return 0;
    if (s.indexOf(',') >= 0) {
      // pt-BR: ponto = milhar, vírgula = decimal.
      s = s.replace(/\./g, '').replace(',', '.');
    }
    // (se só tem ponto, assume ponto decimal — formato US/numérico)
    let n = parseFloat(s);
    if (!isFinite(n)) return 0;
    return neg ? -Math.abs(n) : n;
  }

  // Normaliza texto p/ comparação: sem acento, maiúsculas, espaços colapsados.
  function norm(s) {
    return String(s == null ? '' : s)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().replace(/\s+/g, ' ').trim();
  }

  // Remove sufixos societários e pontuação p/ casar nomes de empresa.
  function stripCompany(s) {
    let n = norm(s);
    n = n.replace(/\b(LTDA|EIRELI|EPP|ME|MEI|S\.?\s?A\.?|S\/A|S\.?C\.?I\.?)\b/g, ' ');
    n = n.replace(/[.\,\/\-]/g, ' ').replace(/\s+/g, ' ').trim();
    return n;
  }

  // Só dígitos (p/ CNPJ).
  function digits(s) { return String(s == null ? '' : s).replace(/\D/g, ''); }

  // Similaridade por tokens significativos (Jaccard sobre o menor conjunto).
  function tokenSim(a, b) {
    const A = new Set(stripCompany(a).split(' ').filter(x => x.length > 2));
    const B = new Set(stripCompany(b).split(' ').filter(x => x.length > 2));
    if (!A.size || !B.size) return 0;
    let inter = 0; A.forEach(x => { if (B.has(x)) inter++; });
    return inter / Math.min(A.size, B.size);
  }

  // Verifica se o nome da empresa "needle" aparece em "hay" (conta analítica).
  function nameAppears(hay, companyName) {
    const h = stripCompany(hay);
    const toks = stripCompany(companyName).split(' ').filter(x => x.length > 3);
    if (!toks.length) return 0;
    let hit = 0;
    toks.forEach(t => { if (h.indexOf(t) >= 0) hit++; });
    return hit / toks.length; // proporção de tokens do nome encontrados
  }

  // Conta espaços iniciais (nível de indentação no balanço).
  function leadSpaces(s) {
    const m = String(s == null ? '' : s).match(/^([ \t ]*)/);
    return m ? m[1].replace(/\t/g, '   ').length : 0;
  }

  // Chave canônica de linha da DRE: normaliza e remove marcadores "(=)", "(-)", "(+-)".
  function dreKey(s) {
    return norm(s).replace(/^\([=+\/\- ]*\)\s*/, '').replace(/\s+/g, ' ').trim();
  }

  // Formata número como pt-BR (apenas p/ exibição na tela).
  function fmt(n) {
    if (n === null || n === undefined || isNaN(n)) return '-';
    return Number(n).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  global.U = { parseNum, norm, stripCompany, digits, tokenSim, nameAppears, leadSpaces, fmt, dreKey };
})(typeof window !== 'undefined' ? window : globalThis);
