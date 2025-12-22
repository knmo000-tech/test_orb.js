// src/export/exportOrbAlignedWithMerges.js

/**
 * orb の表示レイアウトに近い形で .xlsx 出力する
 * - merges(rowspan/colspan) 再現
 * - rowHeaders 省略によるズレを occ + nextFreeCol で解消
 * - numFmt は pivotSpec から決定（Web/Excel を単一定義化）
 *
 * @param {any} pgridwidget
 * @param {string} fileName
 * @param {object|null} pivotSpec  pivotSpec.js の pivotSpec を渡す（任意）
 * @param {object} opts            追加オプション
 *   - XLSX: xlsx-js-style を直接渡したい場合
 *   - defaultNumFmt: マッチしない場合の既定フォーマット（例 "#,##0"）
 */
function exportOrbAlignedWithMerges(
    pgridwidget,
    fileName = "pivot_aligned.xlsx",
    pivotSpec = null,
    opts = {}
) {
    const XLSX = opts.XLSX ?? globalThis.XLSX;
    if (!XLSX) throw new Error("XLSX(xlsx-js-style) が見つかりません。globalThis.XLSX を確認してください。");

    const rowHeaders = pgridwidget?.rows?.headers || [];
    const dataRows = pgridwidget?.dataRows || [];
    const colHeaders = pgridwidget?.columns?.headers || pgridwidget?.cols?.headers || [];

    const bodyRows = Math.max(rowHeaders.length, dataRows.length);
    const colHeaderHeight = colHeaders.length;

    // data列数（右側の値セル数）
    const dataCols = dataRows.reduce((m, r) => Math.max(m, (r || []).length), 0);

    // row header 幅（左側のヘッダ列数）は「hspanの合計の最大」で決める
    const rowHeaderWidth = rowHeaders.reduce((m, r) => {
        const w = (r || []).reduce((s, h) => s + (h?.hspan ? h.hspan(true) : 1), 0);
        return Math.max(m, w);
    }, 0);

    const totalRows = colHeaderHeight + bodyRows;
    const totalCols = rowHeaderWidth + dataCols;

    // ===== PivotSpec 由来の numFmt 引き当て =====
    const fmtMapPack = pivotSpec ? buildExcelNumFmtMap(pivotSpec) : null;
    const defaultNumFmt = opts.defaultNumFmt ?? "#,##0";

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

    function pickMeasureCaptionFromRowHeaders(rhRow) {
        if (!fmtMapPack) return null;
        const { map, norm } = fmtMapPack;

        // rowHeaders 内に measure caption が入っている前提で拾う（完全一致 or 含む）
        for (const h of rhRow || []) {
            const t = norm(toText(h?.value));
            if (!t) continue;

            if (map.has(t)) return t;

            // 「利益率(%)」みたいに装飾されても拾えるように includes でも探す
            for (const key of map.keys()) {
                if (t.includes(key)) return key;
            }
        }
        return null;
    }

    function numFmtForRow(rhRow) {
        if (!fmtMapPack) return null;
        const { map } = fmtMapPack;
        const cap = pickMeasureCaptionFromRowHeaders(rhRow);
        return cap ? (map.get(cap) ?? null) : null;
    }

    // --- blue theme（固定） ---
    const BLUE_HEADER_FILL = "FFEAF8FF"; // 薄い水色
    const BLUE_BORDER = "FF9DD4E8"; // 罫線の水色
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

    const valueStyle = {
        alignment: { vertical: "center", horizontal: "right" },
        border: { top: GRID_BORDER, bottom: GRID_BORDER, left: GRID_BORDER, right: GRID_BORDER },
    };

    // シート用のグリッド（空で埋める）
    const aoa = Array.from({ length: totalRows }, () => Array.from({ length: totalCols }, () => ""));

    // どこが埋まっているか（rowspan/colspanを扱うため）
    const occ = Array.from({ length: totalRows }, () => Array.from({ length: totalCols }, () => false));

    const merges = [];
    function occupy(r, c, rs, cs) {
        const rEnd = Math.min(totalRows, r + rs);
        const cEnd = Math.min(totalCols, c + cs);

        for (let rr = r; rr < rEnd; rr++) {
            for (let cc = c; cc < cEnd; cc++) occ[rr][cc] = true;
        }
        if (rs > 1 || cs > 1) {
            merges.push({ s: { r, c }, e: { r: r + rs - 1, c: c + cs - 1 } });
        }
    }
    function nextFreeCol(r, startC) {
        let c = startC;
        while (c < totalCols && occ[r][c]) c++;
        return c;
    }

    // --- 列ヘッダ（上側）を配置：右側(rowHeaderWidth以降)に置く ---
    for (let r = 0; r < colHeaderHeight; r++) {
        const headers = colHeaders[r] || [];
        let c = rowHeaderWidth;

        for (const h of headers) {
            c = nextFreeCol(r, c);
            const cs = h?.hspan ? h.hspan(true) : 1;
            const rs = h?.vspan ? h.vspan(true) : 1;

            aoa[r][c] = { v: toText(h?.value), t: "s", s: colHeaderStyle };
            occupy(r, c, rs, cs);

            c += cs;
        }
    }

    const RED_FONT = { color: { rgb: "FFFF0000" } }; // 赤
    function applyNegativeRed(style, n) {
        if (typeof n === "number" && n < 0) {
            style.font = { ...(style.font || {}), ...RED_FONT };
        }
        return style;
    }

    // --- 行ヘッダ（左側）を配置 + dataRows を配置 ---
    for (let i = 0; i < bodyRows; i++) {
        const sheetRow = colHeaderHeight + i;
        const headers = rowHeaders[i] || [];
        let c = 0;

        for (const h of headers) {
            c = nextFreeCol(sheetRow, c);
            const cs = h?.hspan ? h.hspan(true) : 1;
            const rs = h?.vspan ? h.vspan(true) : 1;

            aoa[sheetRow][c] = { v: toText(h?.value), t: "s", s: rowHeaderStyle };
            occupy(sheetRow, c, rs, cs);

            c += cs;
            if (c >= rowHeaderWidth) break;
        }

        // ★ pivotSpec から numFmt を決める（見つからなければ default）
        const rowNumFmt = numFmtForRow(rowHeaders[i]) ?? defaultNumFmt;

        const dr = dataRows[i] || [];
        for (let j = 0; j < dr.length; j++) {
            const v = dr[j]?.value;
            const col = rowHeaderWidth + j;

            if (v === null || v === undefined || v === "") {
                aoa[sheetRow][col] = { v: "", t: "s", s: valueStyle };
                continue;
            }

            const n = Number(v);
            if (!Number.isFinite(n)) {
                aoa[sheetRow][col] = { v: String(v), t: "s", s: valueStyle };
                continue;
            }

            const s = applyNegativeRed({ ...valueStyle, numFmt: rowNumFmt }, n);
            aoa[sheetRow][col] = { v: n, t: "n", s };
        }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = merges;

    // ===== 罫線（格子） + 列幅オート調整 =====

    const GRID_LINE = { style: "thin", color: { rgb: "FF9DD4E8" } }; // 青系
    const GRID_BORDER_ALL = { top: GRID_LINE, bottom: GRID_LINE, left: GRID_LINE, right: GRID_LINE };

    // ws の全セルに「格子罫線」を適用（空セルも作って罫線を入れる）
    function applyGridBorders(ws) {
        const ref = ws["!ref"];
        if (!ref) return;

        const range = XLSX.utils.decode_range(ref);
        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const addr = XLSX.utils.encode_cell({ r, c });
                let cell = ws[addr];
                if (!cell) {
                    cell = ws[addr] = { t: "s", v: "" };
                }
                cell.s = cell.s || {};
                cell.s.border = GRID_BORDER_ALL;
            }
        }
    }

    // 文字幅のざっくり計算（日本語は2、ASCIIは1くらいで見積もる）
    function calcTextWidth(str) {
        const s = String(str ?? "");
        let w = 0;
        for (const ch of s) w += ch.charCodeAt(0) <= 0x7f ? 1 : 2;
        return w;
    }

    function countTrailingCommas(fmt) {
        const m = String(fmt ?? "").match(/,+$/);
        return m ? m[0].length : 0;
    }

    // セルの「表示されそうな文字列」をざっくり作る（主要numFmtだけ）
    function displayText(cell) {
        if (!cell) return "";
        const v = cell.v;

        if (cell.t === "n") {
            const n0 = Number(v);
            if (!Number.isFinite(n0)) return "";

            const fmt = cell.s?.numFmt;

            if (fmt && fmt.includes("%")) {
                const decimals = (String(fmt).match(/0\.([0]+)/)?.[1]?.length) ?? 0;
                return `${(n0 * 100).toFixed(decimals)}%`;
            }

            // "#,##0," のような末尾カンマは 1000^k で表示が縮む
            const k = countTrailingCommas(fmt);
            const n = k > 0 ? n0 / Math.pow(1000, k) : n0;

            if (fmt === "#,##0" || (fmt && fmt.startsWith("#,##0"))) {
                return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(n);
            }

            return String(n);
        }

        return String(v ?? "");
    }

    // 列幅をオート調整（上限 maxWidth）
    function autofitColumns(ws, { minWidth = 8, maxWidth = 30, padding = 2 } = {}) {
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

    applyGridBorders(ws);
    autofitColumns(ws, { minWidth: 8, maxWidth: 32, padding: 2 });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pivot");
    XLSX.writeFile(wb, fileName);
}
