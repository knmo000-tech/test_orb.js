// pivotSpec.js（classic script）
(function () {
    // Web表示用：カンマ、％、素（何もしない）
    const fmt = {
        comma: (v) => {
            const n = Number(v);
            if (!Number.isFinite(n)) return "0";
            return String(parseInt(n)).replace(/(\d)(?=(\d\d\d)+(?!\d))/g, "$1,");
        },
        percent1: (v) => {
            const n = Number(v);
            if (!Number.isFinite(n)) return "0.0%";
            return (n * 100).toFixed(1) + "%";
        },
        raw: (v) => {
            const n = Number(v);
            return Number.isFinite(n) ? String(n) : (v == null ? "" : String(v));
        },
    };

    // ★単一定義：ここだけ触れば Web も Excel も揃う
    const pivotSpec = {
        dims: [
            { name: 0, caption: "ベースソース" },
            { name: 1, caption: "ジャンル" },
            { name: 2, caption: "カレーメニュー" },
        ],
        measures: [
            { name: 3, caption: "エネルギー(kcal)", agg: "avg", web: fmt.comma, excel: "#,##0" },
            { name: 4, caption: "売上", agg: "sum", web: fmt.percent1, excel: "0.0%" },
            // { name: 4, caption: "売上", agg: "sum", web: fmt.comma, excel: "#,##0" },
            { name: 5, caption: "利益率", agg: "sum", web: fmt.raw, excel: "General" },
            // { name: 5, caption: "利益率", agg: "sum", web: fmt.percent1, excel: "0.0%" },

            // ✅ 3桁カンマも％もしない「素の数値」
            // { name: 8, caption: "素の数値", agg: "avg", web: fmt.raw, excel: "General" },
        ],
        layout: {
            rows: ["ベースソース", "カレーメニュー"],
            columns: ["ジャンル"],
            data: ["エネルギー(kcal)", "売上", "利益率"],
        },
    };

    function buildOrbConfigFromSpec(spec) {
        const fields = [
            ...spec.dims.map(d => ({ name: d.name, caption: d.caption })),
            ...spec.measures.map(m => ({
                name: m.name,
                caption: m.caption,
                dataSettings: { aggregateFunc: m.agg, formatFunc: m.web },
            })),
        ];
        return { fields, rows: spec.layout.rows, columns: spec.layout.columns, data: spec.layout.data };
    }

    function buildExcelNumFmtMap(spec) {
        const norm = (s) => String(s ?? "").trim();

        const map = new Map(); // ★ Mapにする（has/get/keys が使える）

        for (const m of spec.measures) {
            // measures の excel が "General" / "#,##0" の「文字列」でも、
            // excel:{numFmt:"..."} の「オブジェクト」でも両対応
            const fmt = (m.excel && typeof m.excel === "object")
                ? (m.excel.numFmt ?? null)
                : (m.excel ?? null);

            map.set(norm(m.caption), fmt);
        }

        return { map, norm };
    }

    window.pivotSpec = pivotSpec;
    window.buildOrbConfigFromSpec = buildOrbConfigFromSpec;
    window.buildExcelNumFmtMap = buildExcelNumFmtMap;
})();
