// customexportexcel.js（classic script）
// 依存：window.XLSX（dist/xlsx.bundle.js）
// 依存：window.pivotSpec / window.buildExcelNumFmtMap（pivotSpec.js）

(function () {
  // ====== util ======
  function toText(v) {
    if (v == null) return "";
    if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return String(v);
    if (typeof v === "object") {
      if ("caption" in v && v.caption != null) return String(v.caption);
      if ("name" in v && v.name != null) return String(v.name);
      if ("value" in v && v.value != null) return String(v.value);
    }
    try {
      return JSON.stringify(v);
    } catch {
      return String(v);
    }
  }

  function norm(s) {
    return String(s ?? "").trim();
  }

  // ====== pivotSpec から numFmt を引く仕組み ======
  function getFmtPack(pivotSpec) {
    const spec = pivotSpec || window.pivotSpec;
    if (!spec || typeof window.buildExcelNumFmtMap !== "function") return null;

    const pack = window.buildExcelNumFmtMap(spec); // { map: Map, norm: fn }
    if (!pack || !pack.map) return null;

    return {
      spec,
      map: pack.map,
      norm: typeof pack.norm === "function" ? pack.norm : norm,
    };
  }

  // rowHeaders の1行から「この行がどのメジャーか」を推測して numFmt を返す
  function pickMeasureKeyFromRowHeaders(rowHeaderCells, fmtPack) {
    if (!fmtPack) return null;
    const { map, norm: nrm } = fmtPack;

    // 完全一致（最優先）
    for (const h of rowHeaderCells || []) {
      const t = nrm(toText(h?.value));
      if (t && typeof map.has === "function" && map.has(t)) return t;
    }

    // includes（“利益率(%)”など装飾がある場合の保険）
    if (typeof map.keys === "function") {
      const texts = (rowHeaderCells || []).map((h) => nrm(toText(h?.value))).filter(Boolean);
      for (const key of map.keys()) {
        for (const t of texts) {
          if (t.includes(key)) return key;
        }
      }
    }

    return null;
  }

  function getRowNumFmt(rowHeaderCells, fmtPack, fallback = "General") {
    if (!fmtPack) return fallback;
    const key = pickMeasureKeyFromRowHeaders(rowHeaderCells, fmtPack);
    if (!key) return fallback;

    const fmt = fmtPack.map.get(key);
    return fmt == null ? fallback : fmt; // null/undefined は General に落とす
  }

  // ====== style ======
  const BLUE_HEADER_FILL = "FFEAF8FF";
  const BLUE_BORDER = "FF9DD4E8";
  const GRID_BORDER = { style: "thin", color: { rgb: BLUE_BORDER } };

  const colHeaderStyle = {
    font: { bold: true },
    alignment: { vertical: "center", horizontal: "center", wrapText: true },
    fill: { patternType: "solid", fgColor: { rgb: BLUE_HEADER_FILL } },
    border: { top: GRID_BORDER, bottom: GRID_BORDER, left: GRID_BORDER, right: GRID_BORDER },
  };

  const rowHeaderStyle = {
    font: { bold: true },
    alignment: { vertical: "center", horizontal: "left", wrapText: true },
    fill: { patternType: "solid", fgColor: { rgb: BLUE_HEADER_FILL } },
    border: { top: GRID_BORDER, bottom: GRID_BORDER, left: GRID_BORDER, right: GRID_BORDER },
  };

  const valueStyleBase = {
    alignment: { vertical: "center", horizontal: "right" },
    border: { top: GRID_BORDER, bottom: GRID_BORDER, left: GRID_BORDER, right: GRID_BORDER },
  };

  const RED_FONT = { color: { rgb: "FFFF0000" } };

  function styleWithNegativeRed(baseStyle, n) {
    const s = { ...(baseStyle || {}) };
    if (typeof n === "number" && n < 0) {
      s.font = { ...(s.font || {}), ...RED_FONT };
    }
    return s;
  }

  // ====== grid/merge helpers ======
  function build2D(rows, cols, init) {
    return Array.from({ length: rows }, () => Array.from({ length: cols }, () => init));
  }

  function countSpan(h, kind) {
    const fn = kind === "h" ? h?.hspan : h?.vspan;
    return typeof fn === "function" ? fn(true) : 1;
  }

  function nextFreeCol(occRow, start) {
    let c = start;
    while (c < occRow.length && occRow[c]) c++;
    return c;
  }

  // ====== borders + autofit（簡易）=====
  function applyGridBorders(XLSX, ws) {
    const ref = ws["!ref"];
    if (!ref) return;
    const range = XLSX.utils.decode_range(ref);

    const GRID_LINE = { style: "thin", color: { rgb: BLUE_BORDER } };
    const BORDER_ALL = { top: GRID_LINE, bottom: GRID_LINE, left: GRID_LINE, right: GRID_LINE };

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        let cell = ws[addr];
        if (!cell) cell = ws[addr] = { t: "s", v: "" };
        cell.s = cell.s || {};
        cell.s.border = BORDER_ALL;
      }
    }
  }

  function calcTextWidth(str) {
    const s = String(str ?? "");
    let w = 0;
    for (const ch of s) w += ch.charCodeAt(0) <= 0x7f ? 1 : 2;
    return w;
  }

  function displayText(cell) {
    if (!cell) return "";
    if (cell.t === "n") {
      const n = Number(cell.v);
      if (!Number.isFinite(n)) return "";
      // ざっくり。%だけ考慮
      const fmt = cell.z || cell.s?.numFmt;
      if (fmt && String(fmt).includes("%")) {
        const decimals = (String(fmt).match(/0\.([0]+)/)?.[1]?.length) ?? 0;
        return `${(n * 100).toFixed(decimals)}%`;
      }
      return String(n);
    }
    return String(cell.v ?? "");
  }

  function autofitColumns(XLSX, ws, { minWidth = 8, maxWidth = 32, padding = 2 } = {}) {
    const ref = ws["!ref"];
    if (!ref) return;
    const range = XLSX.utils.decode_range(ref);
    const colCount = range.e.c - range.s.c + 1;
    const widths = Array(colCount).fill(minWidth);

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        const text = displayText(cell);
        const w = Math.min(maxWidth, Math.max(minWidth, calcTextWidth(text) + padding));
        widths[c - range.s.c] = Math.max(widths[c - range.s.c], w);
      }
    }

    ws["!cols"] = widths.map((wch) => ({ wch }));
  }

  // ======================================================================
  // export function (global)
  // ======================================================================
  window.exportOrbAlignedWithMerges = function exportOrbAlignedWithMerges(
    pgridwidget,
    fileName = "pivot.xlsx",
    pivotSpec // 省略可（省略時は window.pivotSpec を使う）
  ) {
    const XLSX = window.XLSX;
    if (!XLSX) throw new Error("XLSX(xlsx-js-style) が見つかりません。dist/xlsx.bundle.js を確認してください。");

    const fmtPack = getFmtPack(pivotSpec);

    const rowHeaders = pgridwidget?.rows?.headers || [];
    const dataRows = pgridwidget?.dataRows || [];
    const colHeaders = pgridwidget?.columns?.headers || pgridwidget?.cols?.headers || [];

    const bodyRows = Math.max(rowHeaders.length, dataRows.length);
    const colHeaderHeight = colHeaders.length;

    const dataCols = dataRows.reduce((m, r) => Math.max(m, (r || []).length), 0);

    const rowHeaderWidth = rowHeaders.reduce((m, r) => {
      const w = (r || []).reduce((s, h) => s + countSpan(h, "h"), 0);
      return Math.max(m, w);
    }, 0);

    const totalRows = colHeaderHeight + bodyRows;
    const totalCols = rowHeaderWidth + dataCols;

    const aoa = build2D(totalRows, totalCols, "");
    const occ = build2D(totalRows, totalCols, false);
    const merges = [];

    function occupy(r, c, rs, cs) {
      const rEnd = Math.min(totalRows, r + rs);
      const cEnd = Math.min(totalCols, c + cs);
      for (let rr = r; rr < rEnd; rr++) {
        for (let cc = c; cc < cEnd; cc++) occ[rr][cc] = true;
      }
      if (rs > 1 || cs > 1) merges.push({ s: { r, c }, e: { r: r + rs - 1, c: c + cs - 1 } });
    }

    // ===== columns headers（上側）=====
    for (let r = 0; r < colHeaderHeight; r++) {
      const headers = colHeaders[r] || [];
      let c = rowHeaderWidth;

      for (const h of headers) {
        c = nextFreeCol(occ[r], c);
        const cs = countSpan(h, "h");
        const rs = countSpan(h, "v");
        aoa[r][c] = { v: toText(h?.value), t: "s", s: colHeaderStyle };
        occupy(r, c, rs, cs);
        c += cs;
      }
    }

    // ===== rows headers + values（本体）=====
    for (let i = 0; i < bodyRows; i++) {
      const sheetRow = colHeaderHeight + i;

      // row headers（左側）
      const headers = rowHeaders[i] || [];
      let c = 0;
      for (const h of headers) {
        c = nextFreeCol(occ[sheetRow], c);
        const cs = countSpan(h, "h");
        const rs = countSpan(h, "v");
        aoa[sheetRow][c] = { v: toText(h?.value), t: "s", s: rowHeaderStyle };
        occupy(sheetRow, c, rs, cs);
        c += cs;
        if (c >= rowHeaderWidth) break;
      }

      // ★ pivotSpec からこの行の numFmt を決める
      const rowNumFmt = getRowNumFmt(rowHeaders[i], fmtPack, "General");

      // values（右側）
      const dr = dataRows[i] || [];
      for (let j = 0; j < dr.length; j++) {
        const col = rowHeaderWidth + j;
        const v = dr[j]?.value;

        if (v === null || v === undefined || v === "") {
          aoa[sheetRow][col] = { v: "", t: "s", s: valueStyleBase };
          continue;
        }

        const n = Number(v);
        if (!Number.isFinite(n)) {
          aoa[sheetRow][col] = { v: String(v), t: "s", s: valueStyleBase };
          continue;
        }

        const s = styleWithNegativeRed(valueStyleBase, n);

        // ★ numFmt は spec に従う（% / カンマ / General など）
        // xlsx-js-style は `cell.z` が効きやすいので z に入れ、念のため s.numFmt も入れる
        if (rowNumFmt) {
          s.numFmt = rowNumFmt;
        }

        aoa[sheetRow][col] = { v: n, t: "n", z: rowNumFmt || "General", s };
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = merges;

    applyGridBorders(XLSX, ws);
    autofitColumns(XLSX, ws, { minWidth: 8, maxWidth: 32, padding: 2 });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pivot");
    XLSX.writeFile(wb, fileName);
  };
})();
