window.onload = () => CreatePivotTable(demo_data)

const CreatePivotTable = (data) => {
    const config = {
        dataSource: data,
        canMoveFields: true,
        dataHeadersLocation: 'row',
        width: 1906,
        height: 900,
        theme: 'blue', /*テーマカラーを指定できます。用意されてるテーマは
    bootstrap、red、blue、green、orange、flower、gray、black、whiteの9色*/
        toolbar: {
            visible: true,
            exportToExcel: (pgridComponent) => {
                // ★あなたの関数だけ呼ぶ（デフォルトxlsは出なくなる）
                exportOrbAlignedWithMerges(pgridComponent.props.pgridwidget, "pivot.xlsx");
            },
        },
        grandTotal: { //このあたりは全て初期値の設定です。ツールバーのvisibleをtrueにしていれば後から変更できます。
            rowsvisible: true,
            columnsvisible: false
        },
        subTotal: {
            visible: true,
            collapsed: true,
            collapsible: true
        },
        rowSettings: {
            subTotal: {
                visible: true,
                collapsed: true,
                collapsible: true
            }
        },
        columnSettings: {
            subTotal: {
                visible: true,
                collapsed: true,
                collapsible: true
            }
        },
        fields: [ //テーブルで扱う項目を作成します。
            {
                name: 0, /*このデータに該当するdataSourceの列番号を指定します。
        プロパティ名おかしいんですが多分{"ベースソース":ポークカレー}みたいなJSONをデータソースにすると
        ここが"ベースソース"になるんじゃないかなあと予想してます。未検証*/
                caption: 'ベースソース'
            },
            {
                name: 1,
                caption: 'ジャンル'
            },
            {
                name: 2,
                caption: 'カレーメニュー'
            },
            {
                name: 3,
                caption: 'エネルギー(kcal)',
                dataSettings: {
                    aggregateFunc: 'avg',
                    /*この項目を集計対象にした際に使用する関数です。
                    count、sum、min、max、avg、prod、stdev、stdevp、var、varpの10種類から選びます。
                   デフォルト値はsum*/
                    formatFunc: val => val ? String(parseInt(val)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,') : 0
                    //format関数みたいなやつです。この例では数値を切り捨ててカンマ区切りを入れています。
                }
            },
            {
                name: 4,
                caption: '売上',
                dataSettings: {
                    aggregateFunc: 'avg',
                    formatFunc: val => val ? String(parseInt(val)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,') : 0
                }
            },
            {
                name: 5,
                caption: '利益率',
                dataSettings: {
                    aggregateFunc: 'avg',
                    //   formatFunc: val => val ? String(parseInt(val)).replace( /(\d)(?=(\d\d\d)+(?!\d))/g, '$1,') : 0
                    formatFunc: (val) => {
                        const n = Number(val);
                        if (!Number.isFinite(n)) return '0%';
                        return `${(n * 100).toFixed(1)}%`; // 小数1桁。不要なら toFixed(0)
                    }
                }
            },
            {
                name: 6,
                caption: '炭水化物(g)',
                dataSettings: {
                    aggregateFunc: 'avg',
                    formatFunc: val => val ? String(parseInt(val)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,') : 0
                }
            },
            {
                name: 7,
                caption: '塩分(g)',
                dataSettings: {
                    aggregateFunc: 'avg',
                    formatFunc: val => val ? String(parseInt(val)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, '$1,') : 0
                }
            }
        ],
        //初期配置
        rows: ['ベースソース', 'カレーメニュー'],
        columns: ['ジャンル'],
        data: ['エネルギー(kcal)', '売上', '利益率'],
    }
    const pgridwidget = new orb.pgridwidget(config)
    pgridwidget.render(document.getElementById('orb'))
    //demo.htmlでdiv要素に指定したID

    // exportOrbAlignedWithMerges(pgridwidget, "pivot_aligned.xlsx"); (pgridwidget);
}


// /**
//  * orb の pgridwidget から XLSX を生成して保存
//  */
// function exportOrbWithXlsxJsStyle(pgridwidget, fileName = "pivot.xlsx") {
//   const config =
//     pgridwidget?.pgrid?.config ||
//     pgridwidget?.config ||
//     pgridwidget?.pgridwidget?.config;

//   if (!config) throw new Error("config not found on pgridwidget");

//   const dataFields = config.dataFields || []; // orb 側が生成している dataFields を期待
//   const dataHeadersInRows = config.dataHeadersLocation === "rows";

//   // orb で確実に使えていたもの（あなたの export 調査の通り）
//   const rowHeaders = pgridwidget.rows?.headers || [];
//   const dataRows = pgridwidget.dataRows || [];

//   // columns ヘッダは環境差があり得るので両対応
//   const colHeaders =
//     pgridwidget.columns?.headers ||
//     pgridwidget.cols?.headers ||
//     [];

//   // ---- スタイル定義（例） ----
//   const headerStyle = {
//     font: { bold: true },
//     alignment: { wrapText: true, vertical: "center" },
//     fill: { patternType: "solid", fgColor: { rgb: "FFEFEFEF" } },
//     border: {
//       top:    { style: "thin", color: { rgb: "FFCCCCCC" } },
//       bottom: { style: "thin", color: { rgb: "FFCCCCCC" } },
//       left:   { style: "thin", color: { rgb: "FFCCCCCC" } },
//       right:  { style: "thin", color: { rgb: "FFCCCCCC" } },
//     },
//   };

//   const negativeNumberStyle = {
//     font: { color: { rgb: "FFFF0000" } }, // 赤 
//   };

//   // dataField ごとの Excel 表示フォーマットを決める（例）
//   // ここはあなたのフィールド定義に合わせて増やす
//   function numFmtForField(field) {
//     // 例: caption/name で判定してもいいし、field に独自メタを持たせてもOK
//     const key = (field && (field.name || field.caption)) || "";
//     if (/率|%|percent|ratio/i.test(key)) return "0.0%";     // パーセント 
//     if (/金額|amount|yen|円/i.test(key)) return "#,##0";    // 3桁カンマ
//     return null;
//   }

//   // 行ヘッダから「この行はどのメジャー（dataField）か」を決める（dataHeadersLocation='rows' 用）
// // 行ヘッダから「この行はどのメジャー（dataField）か」を決める（uiheaders不要版）
// function findRowDataFieldIndex(currRow) {
//   // dataFields の name/caption を検索しやすい形に
//   const keys = dataFields.map((df, idx) => ({
//     idx,
//     name: df?.name,
//     caption: df?.caption,
//   }));

//   for (const h of currRow) {
//     const hv = h?.value;

//     // 1) hv が {name, caption} を持つオブジェクトの場合
//     if (hv && typeof hv === "object") {
//       const key = hv.name ?? hv.caption;
//       if (key != null) {
//         const hit = keys.find(k => k.name === key || k.caption === key);
//         if (hit) return hit.idx;
//       }
//     }

//     // 2) hv が文字列の場合（画面上の表示値がそのまま入ってるケース）
//     if (typeof hv === "string") {
//       const hit = keys.find(k => k.name === hv || k.caption === hv);
//       if (hit) return hit.idx;
//     }
//   }

//   return 0; // 見つからなければ 0 にフォールバック
// }


//   // セル作成ユーティリティ（xlsx-js-style は {v,t,s} を aoa_to_sheet で受けられる） 
//   function sCell(text) {
//     return { v: text ?? "", t: "s", s: headerStyle };
//   }
//   function nCell(num, style, fmt) {
//     const c = { v: num ?? null, t: "n" };
//     if (style || fmt) c.s = { ...(style || {}), ...(fmt ? { numFmt: fmt } : {}) };
//     return c;
//   }

//   // ---- まずは「最小で値が合う」2D配列（AOA）を作る ----
//   // 1) 列ヘッダは一旦 “フラット” に 1 行だけ出す（後で段数対応/結合を足せる）
//   //    orb の colHeaders 構造が複雑なら、ここはまず空でもOK
//   const aoa = [];

//   if (colHeaders.length > 0) {
//     // 最小：最上段だけ出す
//     const top = colHeaders[0] || [];
//     aoa.push(top.map(h => sCell(h?.value ?? "")));
//   }

//   // 2) 本体：各行 = row headers（フラット） + data cells
//   for (let i = 0; i < rowHeaders.length; i++) {
//     const currRow = rowHeaders[i] || [];
//     const row = [];

//     // row headers（最小：並べるだけ。結合やspanは後で）
//     for (const h of currRow) row.push(sCell(h?.value ?? ""));

//     const dr = dataRows[i] || [];

//     // この行のメジャー（rowsにdataHeadersが来る構成の場合）
//     const rowDfIndex = dataHeadersInRows ? findRowDataFieldIndex(currRow) : null;

//     for (let j = 0; j < dr.length; j++) {
//       const raw = dr[j]?.value;

//       // どの dataField のセルか
//       const dfIndex = dataHeadersInRows
//         ? rowDfIndex
//         : (dataFields.length ? (j % dataFields.length) : 0);

//       const field = dataFields[dfIndex];
//       const fmt = numFmtForField(field);

//       // 数値として入れる（Excelに計算させたいので、表示用文字列にはしない）
//       const num = raw == null || raw === "" ? null : Number(raw);

//       // マイナスだけ赤（プラスは黒＝未指定でOK）
//       const st = (typeof num === "number" && num < 0) ? negativeNumberStyle : null;

//       // “null” は空セルにしたいので sCell ではなく未設定に近い nCell
//       row.push(num == null ? { v: "", t: "s" } : nCell(num, st, fmt));
//     }

//     aoa.push(row);
//   }

//   // ---- Sheet / Workbook 化 ----
//   const ws = XLSX.utils.aoa_to_sheet(aoa);                 // 
//   const wb = XLSX.utils.book_new();                        // 
//   XLSX.utils.book_append_sheet(wb, ws, "Pivot");           // 
//   XLSX.writeFile(wb, fileName);                            // 
// }

function exportOrbAlignedWithMerges(pgridwidget, fileName = "pivot_aligned.xlsx") {
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

    function toText(v) {
        if (v == null) return "";
        if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return String(v);
        if (typeof v === "object") {
            if ("caption" in v && v.caption != null) return String(v.caption);
            if ("name" in v && v.name != null) return String(v.name);
            if ("value" in v && v.value != null) return String(v.value);
        }
        try { return JSON.stringify(v); } catch { return String(v); }
    }
    // --- blue theme（固定） ---
    const BLUE_HEADER_FILL = "FFEAF8FF";   // 薄い水色
    const BLUE_BORDER = "FF9DD4E8";   // 罫線の水色
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
    const aoa = Array.from({ length: totalRows }, () =>
        Array.from({ length: totalCols }, () => "")
    );

    // どこが埋まっているか（rowspan/colspanを扱うため）
    const occ = Array.from({ length: totalRows }, () =>
        Array.from({ length: totalCols }, () => false)
    );

    const merges = [];
    function occupy(r, c, rs, cs) {
        for (let rr = r; rr < r + rs; rr++) {
            for (let cc = c; cc < c + cs; cc++) occ[rr][cc] = true;
        }
        if (rs > 1 || cs > 1) merges.push({ s: { r, c }, e: { r: r + rs - 1, c: c + cs - 1 } });
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

    // --- 行ヘッダ（左側）を配置：vspan/hspanを見てズレないように置く ---
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

        function isPercentRow(rhRow) {
            return (rhRow || []).some(h => /利益率|粗利率/.test(toText(h?.value)));
        }

        const dr = dataRows[i] || [];
        const percentRow = isPercentRow(rowHeaders[i]);

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

            if (percentRow) {
                // 利益率 / 粗利率：% 表示（値はそのまま）
                aoa[sheetRow][col] = {
                    v: n,
                    t: "n",
                    s: { ...valueStyle, numFmt: "0.0%" }
                };
            } else {
                // それ以外：3桁カンマ（通常の数値）
                aoa[sheetRow][col] = {
                    v: n,
                    t: "n",
                    s: { ...valueStyle, numFmt: "#,##0" }
                };

                /*
                // --- 1000円単位にしたい場合（復活用） ---
                // 値を /1000 して「千円」単位で表示する
                aoa[sheetRow][col] = {
                  v: n / 1000,
                  t: "n",
                  s: { ...valueStyle, numFmt: "#,##0" } // 見た目はカンマ区切り
                  // 例：単位も表示したいなら → numFmt: '#,##0"千円"'
                };
                */
            }
        }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = merges;

    // --- 追加：利益率の行を % 表示にする（Excel 側の表示形式だけ変える） ---
    function rowHasText(rhRow, re) {
        return (rhRow || []).some(h => re.test(toText(h?.value)));
    }

    for (let i = 0; i < bodyRows; i++) {
        // 「利益率」の行だけ対象（必要なら正規表現を調整）
        if (!rowHasText(rowHeaders[i], /利益率/)) continue;

        for (let j = 0; j < dataCols; j++) {
            const addr = XLSX.utils.encode_cell({
                r: colHeaderHeight + i,      // データ開始行
                c: rowHeaderWidth + j        // データ開始列
            });

            const cell = ws[addr];
            if (!cell || cell.t !== "n") continue; // 数値セルだけ

            cell.s = cell.s || {};
            cell.s.numFmt = "0.0%";               // 例：3812.0% のように出したいならこれ
        }
    }

    // ===== 追加ここから =====

    // 罫線（格子）
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
                    // 空セルにも罫線を入れたいので作る
                    cell = ws[addr] = { t: "s", v: "" };
                }
                cell.s = cell.s || {};
                // fill/font などは残して border だけ上書き
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

    // セルの「表示されそうな文字列」をざっくり作る（numFmt の主要2種だけ対応）
    function displayText(cell) {
        if (!cell) return "";
        const v = cell.v;

        if (cell.t === "n") {
            const n = Number(v);
            if (!Number.isFinite(n)) return "";

            const fmt = cell.s?.numFmt;

            if (fmt && fmt.includes("%")) {
                // "0.0%" 想定
                const decimals = (fmt.match(/0\.([0]+)/)?.[1]?.length) ?? 0;
                return `${(n * 100).toFixed(decimals)}%`;
            }

            if (fmt === "#,##0") {
                return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(n);
            }

            return String(n);
        }

        // 文字列
        return String(v ?? "");
    }

    // 列幅をオート調整（上限 maxWidth）
    // maxWidth/minWidth は「Excelのwch（だいたい文字数）」のイメージ
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

        ws["!cols"] = widths.map(wch => ({ wch }));
    }

    // 実行
    applyGridBorders(ws);
    autofitColumns(ws, { minWidth: 8, maxWidth: 32, padding: 2 }); // maxWidthは好みで調整

    // ===== 追加ここまで =====


    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pivot");
    XLSX.writeFile(wb, fileName);
}
