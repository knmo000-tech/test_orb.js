window.onload = () => CreatePivotTable(window.demo_data);

function CreatePivotTable(data) {
    const fromSpec = buildOrbConfigFromSpec(pivotSpec);

    const config = {
        dataSource: data,
        canMoveFields: true,
        dataHeadersLocation: "row",
        width: 1906,
        height: 900,
        theme: "blue",
        toolbar: {
            visible: true,
            exportToExcel: (pgridComponent) => {
                // ★ exporter は別ファイル + spec を渡す
                exportOrbAlignedWithMerges(
                    pgridComponent.props.pgridwidget,
                    "pivot.xlsx",
                    pivotSpec
                );
            },
        },
        grandTotal: {
            rowsvisible: true,
            columnsvisible: false,
        },
        subTotal: {
            visible: true,
            collapsed: true,
            collapsible: true,
        },
        rowSettings: {
            subTotal: {
                visible: true,
                collapsed: true,
                collapsible: true,
            },
        },
        columnSettings: {
            subTotal: {
                visible: true,
                collapsed: true,
                collapsible: true,
            },
        },

        // ★ fields/rows/columns/data は spec から一括生成
        fields: fromSpec.fields,
        rows: fromSpec.rows,
        columns: fromSpec.columns,
        data: fromSpec.data,
    };

    const pgridwidget = new orb.pgridwidget(config);
    pgridwidget.render(document.getElementById("orb"));
}