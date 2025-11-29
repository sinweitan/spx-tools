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
        tables.sort((a, b) => b.rows.length - a.rows.length); // biggest table
        return tables[0];
    }

    const table = pickMainTable();
    if (!table) return;

    // 1) table -> temp sheet -> 2D array
    const tmpWs = XLSX.utils.table_to_sheet(table);
    const data = XLSX.utils.sheet_to_json(tmpWs, { header: 1, raw: true });
    if (!data.length) {
        alert("Table appears to be empty.");
        return;
    }

    // 2) remove columns by LETTERS: C,F,G,H,I,K
    // letters -> 0-based indexes on header row: A=0,B=1,C=2,...
    const removeIdxSet = new Set([2, 5, 6, 7, 8, 10]);

    const filtered = data.map((row) =>
        row.filter((_, idx) => !removeIdxSet.has(idx))
    );

    // 3) prepend Code + PIN columns
    const newData = filtered.map((row, rIdx) => {
        if (rIdx === 0) {
            return ["Code", "PIN", ...row]; // header row
        }
        return ["", "", ...row];           // data rows
    });

    // 4) array -> sheet
    const ws = XLSX.utils.aoa_to_sheet(newData);

    // 5) auto column widths
    const numCols = newData[0] ? newData[0].length : 0;
    const colWidths = [];
    for (let C = 0; C < numCols; C++) {
        let maxWidth = 8;
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

    // 6) all row heights ≈ 60px
    ws["!rows"] = new Array(newData.length)
        .fill(null)
        .map(() => ({ hpx: 60 }));

    // 7) thin borders + vertical center + wrap text
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

    // 8) build workbook + download
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SPX");

    function pad(n) { return String(n).padStart(2, "0"); }
    const d = new Date();
    const fname = `SPX_export_${d.getFullYear()}${pad(
        d.getMonth() + 1
    )}${pad(d.getDate())}_${pad(d.getHours())}${pad(
        d.getMinutes()
    )}${pad(d.getSeconds())}.xlsx`;

    XLSX.writeFile(wb, fname);
})();
