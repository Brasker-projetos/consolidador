/* loader.js — produz um grid (array de arrays) a partir de um .xls/.xlsx.
   Para .xlsx e .xls que o SheetJS lê: usa sheet_to_json.
   Para .xls BIFF8 que o SheetJS abre mas NÃO popula (caso dos exports contábeis):
   faz fallback para um leitor BIFF mínimo usando o stream cru (XLSX.CFB) + SST. */
(function (global) {
  'use strict';
  const XLSX = global.XLSX;

  function rkVal(rk) {
    const cents = rk & 1, isInt = rk & 2; let v;
    if (isInt) { v = rk >> 2; }
    else { const b = new ArrayBuffer(8); const d = new DataView(b); d.setUint32(4, (rk & 0xFFFFFFFC) >>> 0, true); v = d.getFloat64(0, true); }
    return cents ? v / 100 : v;
  }

  function readUnicode(stream, off, cch) {
    const grbit = stream[off]; let p = off + 1, str = '';
    if (grbit & 1) { for (let i = 0; i < cch; i++) { str += String.fromCharCode(stream[p] | (stream[p + 1] << 8)); p += 2; } }
    else { for (let i = 0; i < cch; i++) { str += String.fromCharCode(stream[p]); p++; } }
    return str;
  }

  // Reconstrói o grid varrendo os records do substream da planilha (BIFF8).
  function biffGrid(u8, wb) {
    const SST = (wb.Strings || []).map(s => typeof s === 'string' ? s : (s.t || ''));
    const cfb = XLSX.CFB.read(u8, { type: 'array' });
    const entry = cfb.FileIndex.find(f => /workbook|book/i.test(f.name));
    if (!entry) return [];
    const stream = Uint8Array.from(entry.content);
    const dv = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);

    // acha o BOF do worksheet (docType 0x10); se não houver, começa do 0
    let p = 0, wsStart = -1;
    while (p + 4 <= stream.length) {
      const rec = stream[p] | (stream[p + 1] << 8), len = stream[p + 2] | (stream[p + 3] << 8);
      if (rec === 0x809) { const dt = stream[p + 6] | (stream[p + 7] << 8); if (dt === 0x10) { wsStart = p; break; } }
      p += 4 + len;
    }
    if (wsStart < 0) wsStart = 0;

    const grid = [];
    const set = (r, c, v) => { if (!grid[r]) grid[r] = []; grid[r][c] = v; };
    p = wsStart;
    let first = true;
    while (p + 4 <= stream.length) {
      const rec = stream[p] | (stream[p + 1] << 8), len = stream[p + 2] | (stream[p + 3] << 8), d = p + 4;
      if (rec === 0xa && !first) break;            // EOF do substream
      first = false;
      try {
        if (rec === 0xfd) {                          // LABELSST
          const r = stream[d] | (stream[d + 1] << 8), c = stream[d + 2] | (stream[d + 3] << 8), isst = dv.getUint32(d + 6, true);
          set(r, c, SST[isst] != null ? SST[isst] : '');
        } else if (rec === 0x203) {                  // NUMBER
          set(stream[d] | (stream[d + 1] << 8), stream[d + 2] | (stream[d + 3] << 8), dv.getFloat64(d + 6, true));
        } else if (rec === 0x27e) {                  // RK
          set(stream[d] | (stream[d + 1] << 8), stream[d + 2] | (stream[d + 3] << 8), rkVal(dv.getInt32(d + 6, true)));
        } else if (rec === 0xbd) {                   // MULRK
          const r = stream[d] | (stream[d + 1] << 8); let c = stream[d + 2] | (stream[d + 3] << 8); let q = d + 4;
          const colLast = stream[p + 4 + len - 2] | (stream[p + 4 + len - 1] << 8);
          while (c <= colLast) { set(r, c, rkVal(dv.getInt32(q + 2, true))); q += 6; c++; }
        } else if (rec === 0x204) {                  // LABEL (string inline)
          const r = stream[d] | (stream[d + 1] << 8), c = stream[d + 2] | (stream[d + 3] << 8);
          const cch = stream[d + 6] | (stream[d + 7] << 8);
          set(r, c, readUnicode(stream, d + 8, cch));
        } else if (rec === 0x6) {                    // FORMULA (valor numérico em cache)
          const hi = stream[d + 12] | (stream[d + 13] << 8);
          if (hi !== 0xFFFF) set(stream[d] | (stream[d + 1] << 8), stream[d + 2] | (stream[d + 3] << 8), dv.getFloat64(d + 6, true));
        }
      } catch (e) { /* ignora record malformado isolado */ }
      p += 4 + len;
    }
    return normalize(grid);
  }

  // Garante linhas/colunas contíguas (sem buracos undefined) p/ o parser.
  function normalize(grid) {
    const out = [];
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      const nr = [];
      const maxC = row.length;
      for (let c = 0; c < maxC; c++) nr[c] = row[c] === undefined ? '' : row[c];
      out[r] = nr;
    }
    return out;
  }

  // Loader principal.
  function gridFromBuffer(u8) {
    const wb = XLSX.read(u8, { type: 'array', cellText: true, cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    let grid = (ws && ws['!ref'])
      ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true })
      : [];
    if (!grid || !grid.length) grid = biffGrid(u8, wb);   // fallback BIFF
    return grid;
  }

  global.LOAD = { gridFromBuffer };
})(typeof window !== 'undefined' ? window : globalThis);
