(async function () {
    console.log("SPX Tools Export Script v3 Loaded");

    // Load SheetJS
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    document.head.appendChild(script);
    await new Promise((res) => (script.onload = res));

    function pickMainTable() {
        const tables = Array.from(document.querySelectorAll("table"));
        if (!tables.length) {
            alert("No tables found.");
            return null;
        }
        tables.sort((a, b) => b.rows.length - a.rows.length);
        return tables[0];
    }

    const table = pickMainTable();
    if (!table) return;

    // Convert table → array
    const tmp = XLSX.utils.table_to_sheet(table);
    const data = XLSX.utils.sheet_to_json(tmp, { header: 1, raw: true });
    if (!data.length) return;

    // Remove columns C,F,G,H,I,K (0=A)
    const removeIdx = new Set([2, 5, 6, 7, 8, 10]);
    const filtered = data.map(row =>
        row.filter((_, idx) => !removeIdx.has(idx))
    );

    // Add Code + PIN
    const finalData = filtered.map((row, r) =>
        r === 0 ? ["Code", "PIN", ...row] : ["", "", ...row]
    );

    // Convert to worksheet
    const ws = XLSX.utils.aoa_to_sheet(finalData);

    // Auto width (first 2 fixed)
    const colCount = finalData[0].length;
    ws["!cols"] = [];
    for (let c = 0; c < colCount; c++) {
        if (c < 2) {
            ws["!cols"].push({ wch: 12 });
        } else {
            let max = 10;
            for (let r = 0; r < finalData.length; r++) {
                const val = finalData[r][c];
                if (val) max = Math.max(max, String(val).length + 2);
            }
            ws["!cols"].push({ wch: max });
        }
    }

    // Row height
    ws["!rows"] = new Array(finalData.length)
        .fill(null)
        .map(() => ({ hpx: 60 }));

    // --- REAL DATE CONVERSION (BEST METHOD) ---
    function parseDate(str) {
        if (!str) return null;

        // datetime? e.g. 29/11/2025 9:48
        if (/\d{1,2}\/\d{1,2}\/\d{4}\s+\d/.test(str)) {
            const [d, t] = str.split(" ");
            const [dd, mm, yyyy] = d.split("/").map(Number);
            const [hh, mi] = t.split(":").map(Number);
            return new Date(yyyy, mm - 1, dd, hh, mi);
        }

        // date only e.g. 04/12/2025
        if (/\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
            const [dd, mm, yyyy] = str.split("/").map(Number);
            return new Date(yyyy, mm - 1, dd);
        }

        return null;
    }

    // Convert col F (5) to datetime, col G (6) to date
    function convertColumn(colIndex, withTime) {
        const range = XLSX.utils.decode_range(ws["!ref"]);
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            const cellAddr = XLSX.utils.encode_cell({ r, c: colIndex });
            const cell = ws[cellAddr];
            if (!cell) continue;

            const dt = parseDate(String(cell.v));
            if (!dt) continue;

            cell.v = dt;
            cell.t = "d";
            cell.z = withTime ? "dd/mm/yyyy h:mm" : "dd/mm/yyyy";
        }
    }

    convertColumn(5, true);   // F
    convertColumn(6, false);  // G

    // -----------------------------------------

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SPX");

    function pad(n) { return String(n).padStart(2, "0"); }
    const d = new Date();
    const fname =
        `SPX_export_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}.xlsx`;

    XLSX.writeFile(wb, fname);
})();
