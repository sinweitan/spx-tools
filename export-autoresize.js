(async function () {
    // Load SheetJS
    const script = document.createElement("script");
    script.src =
        "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    document.head.appendChild(script);
    await new Promise((res) => (script.onload = res));

    function pickMainTable() {
        const tables = Array.from(document.querySelectorAll("table"));
        if (!tables.length) {
            alert("No tables found on this page.");
            return null;
        }
        // pick the largest (most rows) table
        tables.sort((a, b) => b.rows.length - a.rows.length);
        return tables[0];
    }

    const table = pickMainTable();
    if (!table) return;

    // STEP 1: table -> temp sheet -> 2D array
    const tmpWs = XLSX.utils.table_to_sheet(table);
    const data = XLSX.utils.sheet_to_json(tmpWs, { header: 1, raw: true });

    if (!data.length) {
        alert("Table appears to be empty.");
        return;
    }

    // STEP 2: remove specific columns by index
    // Excel: A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10
    // Remove: C, F, G, H, I, K -> indices: 2, 5, 6, 7, 8, 10
    const removeIdxSet = new Set([2, 5, 6, 7, 8, 10]);

    const filtered = data.map((row) =>
        row.filter((_, idx) => !removeIdxSet.has(idx))
    );

    // STEP 3: prepend 2 columns: Code, PIN
    const newData = filtered.map((row, rIdx) => {
        if (rIdx === 0) {
            // header row
            return ["Code", "PIN", ...row];
        } else {
            // data rows (blank values for Code, PIN)
            return ["", "", ...row];
        }
    });

    // STEP 4: array -> sheet
    const ws = XLSX.utils.aoa_to_sheet(newData);

    // STEP 5: auto column widths (based on newData)
    const numCols = newData[0] ? newData[0].length : 0;
    const colWidths = [];

    for (let C = 0; C < numCols; C++) {
        let maxWidth = 8; // minimum width
        for (let R = 0; R < newData.length; R++) {
            const val = newData[R][C];
            if (val !== undefined && val !== null) {
                const len = String(val).length;
                if (len > maxWidth) maxWidth = len;
            }
        }
        colWidths.push({ wch: maxWidth + 2 });
    }

    ws["!cols"] = colWidths;

    // STEP 6: set all row heights to ~60px
    // 60px ~ 45pt, but SheetJS supports direct pixel height via hpx
    ws["!rows"] = new Array(newData.length)
        .fill(null)
        .map(() => ({ hpx: 60 }));

    // STEP 7: add borders to all cells (and center vertically + wrap text)
    // Note: full styling support depends on Excel + SheetJS; if some
    // apps ignore borders, the data is still all there.
    if (ws["!ref"]) {
        const range = XLSX.utils.decode_range(ws["!ref"]);
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[addr];
                if (!cell) continue;
                cell.s = cell.s || {};
                cell.s.border = {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } },
                };
                cell.s.alignment = {
                    vertical: "center",
                    wrapText: true,
                };
            }
        }
    }

    // STEP 8: build workbook and download
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SPX");

    function pad(n) {
        return String(n).padStart(2, "0");
    }
    const d = new Date();
    const fname = `SPX_export_${d.getFullYear()}${pad(
        d.getMonth() + 1
    )}${pad(d.getDate())}_${pad(d.getHours())}${pad(
        d.getMinutes()
    )}${pad(d.getSeconds())}.xlsx`;

    XLSX.writeFile(wb, fname);
})();
