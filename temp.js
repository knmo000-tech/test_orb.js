// pivotSpec.js（完成版：売上加重平均を含む / classic script）
(function () {
  // =========================
  // 1) 固定順ソート（任意）
  // =========================
  const SORT_LISTS = {
    CURRY_MENU_ORDER: [
      "チキンにこみカレー",
      "ロースカツカレー",
      "ポークカレー",
      // ...必要に応じて追加
    ],
  };

  function makeListComparator(listKey, { fallback = "asc", unknown = "last" } = {}) {
    const list = SORT_LISTS[listKey] || [];
    const index = new Map(list.map((v, i) => [String(v), i]));

    return (a, b) => {
      const sa = String(a ?? "");
      const sb = String(b ?? "");
      const hasA = index.has(sa);
      const hasB = index.has(sb);

      if (hasA && !hasB) return unknown === "first" ? -1 : 1;
      if (!hasA && hasB) return unknown === "first" ? 1 : -1;
      if (hasA && hasB) return index.get(sa) - index.get(sb);

      return fallback === "desc"
        ? sb.localeCompare(sa, "ja")
        : sa.localeCompare(sb, "ja");
    };
  }

  function toOrbSort(sortSpec) {
    if (!sortSpec) return undefined;

    if (sortSpec.order === "asc" || sortSpec.order === "desc") {
      return { order: sortSpec.order, customfunc: null };
    }

    if (typeof sortSpec.listKey === "string") {
      return {
        order: null,
        customfunc: makeListComparator(sortSpec.listKey, {
          fallback: sortSpec.fallback ?? "asc",
          unknown: sortSpec.unknown ?? "last",
        }),
      };
    }

    return undefined;
  }

  // =========================
  // 2) Web表示用 formatFunc
  // =========================
  const fmt = {
    commaInt: (val) => {
      const n = Number(val);
      if (!Number.isFinite(n)) return "0";
      return String(parseInt(n)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, "$1,");
    },
    percent1: (val) => {
      const n = Number(val);
      if (!Number.isFinite(n)) return "0.0%";
      return (n * 100).toFixed(1) + "%";
    },
    raw: (val) => {
      const n = Number(val);
      return Number.isFinite(n) ? String(n) : (val == null ? "" : String(val));
    },
  };

  // =========================
  // 3) カスタム aggregateFunc
  //    - orb の signature:
  //      (datafield, intersection, datasource, rowIndexes, colIndexes) => number
  // =========================

  function getIntersectionIndexes(intersection, datasource) {
    if (intersection === "all") {
      return Array.from({ length: datasource.length }, (_, i) => i);
    }
    return Array.isArray(intersection) ? intersection : [];
  }

  function getValueFromRow(row, fieldIndex) {
    const v = Array.isArray(row) ? row[fieldIndex] : row?.[fieldIndex];
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }

  /**
   * 売上加重平均:
   *   (Σ value * weight) / (Σ weight)
   * @param {number} weightFieldIndex 例: 売上が列4なら 4
   */
  function aggWeightedAvgBy(weightFieldIndex) {
    return function (datafield, intersection, datasource /*, rowIdxs, colIdxs */) {
      const idxs = getIntersectionIndexes(intersection, datasource);

      let sumWV = 0;
      let sumW = 0;

      // datafield は orb から「この measure の name(=列番号)」が渡ってくる想定
      const valueFieldIndex =
        typeof datafield === "number" ? datafield : Number(datafield);

      for (const i of idxs) {
        const row = datasource[i];
        const v = getValueFromRow(row, valueFieldIndex);      // 利益率など
        const w = getValueFromRow(row, weightFieldIndex);     // 売上など
        if (v == null || w == null) continue;
        sumWV += v * w;
        sumW += w;
      }

      return sumW === 0 ? 0 : (sumWV / sumW);
    };
  }

  // =========================
  // 4) PivotSpec 本体
  // =========================
  const pivotSpec = {
    dims: [
      { name: 0, caption: "ベースソース" },
      { name: 1, caption: "ジャンル", sort: { order: "asc" } },
      { name: 2, caption: "カレーメニュー", sort: { listKey: "CURRY_MENU_ORDER" } },
    ],

    measures: [
      {
        name: 3,
        caption: "エネルギー(kcal)",
        agg: "avg",
        web: fmt.commaInt,
        excel: "#,##0",
      },
      {
        name: 4,
        caption: "売上",
        agg: "sum",
        web: fmt.commaInt,
        excel: "#,##0",
      },
      {
        name: 5,
        caption: "利益率",
        // ★ここが本題：売上（列4）で加重平均
        agg: aggWeightedAvgBy(4),
        web: fmt.percent1,
        excel: "0.0%",
      },
      {
        name: 6,
        caption: "炭水化物(g)",
        agg: "avg",
        web: fmt.commaInt,
        excel: "#,##0",
      },
      {
        name: 7,
        caption: "塩分(g)",
        agg: "avg",
        web: fmt.commaInt,
        excel: "#,##0",
      },
    ],

    layout: {
      rows: ["ベースソース", "カレーメニュー"],
      columns: ["ジャンル"],
      data: ["エネルギー(kcal)", "売上", "利益率"],
    },
  };

  // =========================
  // 5) orb config 生成（dataSettings のみ）
  // =========================
  function buildOrbConfigFromSpec(spec) {
    const fields = [
      ...spec.dims.map((d) => ({
        name: d.name,
        caption: d.caption,
        sort: toOrbSort(d.sort),
      })),
      ...spec.measures.map((m) => ({
        name: m.name,
        caption: m.caption,
        dataSettings: {
          aggregateFunc: m.agg,     // ★ string | function 両対応
          formatFunc: m.web,
        },
      })),
    ];

    return {
      fields,
      rows: spec.layout.rows,
      columns: spec.layout.columns,
      data: spec.layout.data,
    };
  }

  // =========================
  // 6) Excel numFmt map（exporter 用）
  // =========================
  function buildExcelNumFmtMap(spec) {
    const norm = (s) => String(s ?? "").trim();
    const map = new Map();
    for (const m of spec.measures) map.set(norm(m.caption), m.excel ?? null);
    return { map, norm };
  }

  // window へ公開
  window.pivotSpec = pivotSpec;
  window.buildOrbConfigFromSpec = buildOrbConfigFromSpec;
  window.buildExcelNumFmtMap = buildExcelNumFmtMap;

  // 任意（デバッグ）
  window.__SORT_LISTS__ = SORT_LISTS;
})();
