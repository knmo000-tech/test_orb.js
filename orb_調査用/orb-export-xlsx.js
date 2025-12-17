/**
 * orb の pgridwidget から XLSX を生成して保存
 */
export function exportOrbWithXlsxJsStyle(pgridwidget, fileName = "pivot.xlsx") {
  const config =
    pgridwidget?.pgrid?.config ||
    pgridwidget?.config ||
    pgridwidget?.pgridwidget?.config;

  if (!config) throw new Error("config not found on pgridwidget");

  const dataFields = config.dataFields || []; // orb 側が生成している dataFields を期待
  const dataHeadersInRows = config.dataHeadersLocation === "rows";

  // orb で確実に使えていたもの（あなたの export 調査の通り）
  const rowHeaders = pgridwidget.rows?.headers || [];
  const dataRows = pgridwidget.dataRows || [];

  // columns ヘッダは環境差があり得るので両対応
  const colHeaders =
    pgridwidget.columns?.headers ||
    pgridwidget.cols?.headers ||
    [];

  // ---- スタイル定義（例） ----
  const headerStyle = {
    font: { bold: true },
    alignment: { wrapText: true, vertical: "center" },
    fill: { patternType: "solid", fgColor: { rgb: "FFEFEFEF" } },
    border: {
      top:    { style: "thin", color: { rgb: "FFCCCCCC" } },
      bottom: { style: "thin", color: { rgb: "FFCCCCCC" } },
      left:   { style: "thin", color: { rgb: "FFCCCCCC" } },
      right:  { style: "thin", color: { rgb: "FFCCCCCC" } },
    },
  };

  const negativeNumberStyle = {
    font: { color: { rgb: "FFFF0000" } }, // 赤 
  };

  // dataField ごとの Excel 表示フォーマットを決める（例）
  // ここはあなたのフィールド定義に合わせて増やす
  function numFmtForField(field) {
    // 例: caption/name で判定してもいいし、field に独自メタを持たせてもOK
    const key = (field && (field.name || field.caption)) || "";
    if (/率|%|percent|ratio/i.test(key)) return "0.0%";     // パーセント 
    if (/金額|amount|yen|円/i.test(key)) return "#,##0";    // 3桁カンマ
    return null;
  }

  // 行ヘッダから「この行はどのメジャー（dataField）か」を決める（dataHeadersLocation='rows' 用）
  function findRowDataFieldIndex(currRow) {
    // currRow の中に DATA_HEADER がある想定（あなたの調査と同じ）
    for (const h of currRow) {
      if (h && h.type === uiheaders.HeaderType.DATA_HEADER) {
        const key = h.value && (h.value.name || h.value.caption) ? (h.value.name || h.value.caption) : h.value;
        const idx = dataFields.findIndex(df => df && (df.name === key || df.caption === key));
        return idx >= 0 ? idx : 0;
      }
    }
    return 0;
  }

  // セル作成ユーティリティ（xlsx-js-style は {v,t,s} を aoa_to_sheet で受けられる） 
  function sCell(text) {
    return { v: text ?? "", t: "s", s: headerStyle };
  }
  function nCell(num, style, fmt) {
    const c = { v: num ?? null, t: "n" };
    if (style || fmt) c.s = { ...(style || {}), ...(fmt ? { numFmt: fmt } : {}) };
    return c;
  }

  // ---- まずは「最小で値が合う」2D配列（AOA）を作る ----
  // 1) 列ヘッダは一旦 “フラット” に 1 行だけ出す（後で段数対応/結合を足せる）
  //    orb の colHeaders 構造が複雑なら、ここはまず空でもOK
  const aoa = [];

  if (colHeaders.length > 0) {
    // 最小：最上段だけ出す
    const top = colHeaders[0] || [];
    aoa.push(top.map(h => sCell(h?.value ?? "")));
  }

  // 2) 本体：各行 = row headers（フラット） + data cells
  for (let i = 0; i < rowHeaders.length; i++) {
    const currRow = rowHeaders[i] || [];
    const row = [];

    // row headers（最小：並べるだけ。結合やspanは後で）
    for (const h of currRow) row.push(sCell(h?.value ?? ""));

    const dr = dataRows[i] || [];

    // この行のメジャー（rowsにdataHeadersが来る構成の場合）
    const rowDfIndex = dataHeadersInRows ? findRowDataFieldIndex(currRow) : null;

    for (let j = 0; j < dr.length; j++) {
      const raw = dr[j]?.value;

      // どの dataField のセルか
      const dfIndex = dataHeadersInRows
        ? rowDfIndex
        : (dataFields.length ? (j % dataFields.length) : 0);

      const field = dataFields[dfIndex];
      const fmt = numFmtForField(field);

      // 数値として入れる（Excelに計算させたいので、表示用文字列にはしない）
      const num = raw == null || raw === "" ? null : Number(raw);

      // マイナスだけ赤（プラスは黒＝未指定でOK）
      const st = (typeof num === "number" && num < 0) ? negativeNumberStyle : null;

      // “null” は空セルにしたいので sCell ではなく未設定に近い nCell
      row.push(num == null ? { v: "", t: "s" } : nCell(num, st, fmt));
    }

    aoa.push(row);
  }

  // ---- Sheet / Workbook 化 ----
  const ws = XLSX.utils.aoa_to_sheet(aoa);                 // 
  const wb = XLSX.utils.book_new();                        // 
  XLSX.utils.book_append_sheet(wb, ws, "Pivot");           // 
  XLSX.writeFile(wb, fileName);                            // 
}
