(async function () {
    console.log("SPX Tools Export Script v3 Loaded");

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
        tables.sort((a, b) => b.rows.length - a.rows.length);
        return tables[0];
    }

    const table = pickMainTable();
    if (!table) return;

    // Convert table → 2D array
    const tmpWs = XLSX.utils.table_to_sheet(table);
    const data = XLSX.utils.sheet_to_json(tmpWs, { header: 1, raw: true });
    if (!data.length) return;

    // Remove unwanted columns C,F,G,H,I,K = 2,5,6,7,8,10
    const removeIdx = new Set([2, 5, 6, 7, 8, 10]);
    const filtered = data.map((row) =>
        row.filter((_, idx) => !removeIdx.has(idx))
    );

    // Add Code + PIN columns
    const finalData = filtered.map((row, r) => {
        if (r === 0) return ["Code", "PIN", ...row];
        return ["", "", ...row];
    });

    const ws = XLSX.utils.aoa_to_sheet(finalData);

    // Auto column width (except first 2)
    const colCount = finalData[0].length;
    ws["!cols"] = [];

    for (let i = 0; i < colCount; i++) {
        if (i < 2) {
            ws["!cols"].push({ wch: 12 });
        } else {
            let max = 10;
            for (let r = 0; r < finalData.length; r++) {
                const val = finalData[r][i];
                if (val !== undefined && val !== null) {
                    max = Math.max(max, String(val).length + 2);
                }
            }
            ws["!cols"].push({ wch: max });
        }
    }

    // Row height = 60px
    ws["!rows"] = new Array(finalData.length)
        .fill(null)
        .map(() => ({ hpx: 60 }));

    // Borders + wrap + center
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
                right: { style: "thin", color: { rgb: "000000" } }
            };
            cell.s.alignment = {
                wrapText: true,
                vertical: "center"
            };
        }
    }

    // ===== DATE FORMATTING =====
    function excelSerialFromDMY(dmyStr, withTime) {
        if (!dmyStr) return null;
        const parts = dmyStr.trim().split(/\s+/);
        const datePart = parts[0];
        const timePart = parts[1] || "";

        const [dd, mm, yyyy] = datePart.split(/[\/\-]/).map(Number);
        if (!dd || !mm || !yyyy) return null;

        let hh = 0, min = 0;
        if (withTime && timePart.includes(":")) {
            const [hStr, mStr] = timePart.split(":");
            hh = Number(hStr) || 0;
            min = Number(mStr) || 0;
        }

        const jsDate = new Date(yyyy, mm - 1, dd, hh, min);
        if (isNaN(jsDate.getTime())) return null;

        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        return (jsDate - excelEpoch) / 86400000;
    }

    function formatColumnAsDate(ws, colIndex, withTime) {
        if (!ws["!ref"]) return;
        const r = XLSX.utils.decode_range(ws["!ref"]);
        for (let R = r.s.r + 1; R <= r.e.r; ++R) {
            const addr = XLSX.utils.encode_cell({ r: R, c: colIndex });
            const cell = ws[addr];
            if (!cell || cell.v == null) continue;

            const text = String(cell.v);
            const serial = excelSerialFromDMY(text, withTime);
            if (serial == null) continue;

            cell.v = serial;
            cell.t = "n";
            cell.z = withTime ? "dd/mm/yyyy h:mm" : "dd/mm/yyyy";
        }
    }

    // Format columns F (index 5) & G (index 6)
    formatColumnAsDate(ws, 5, true);   // F: date-time
    formatColumnAsDate(ws, 6, false);  // G: date only

    // ============================

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SPX");

    function pad(n) { return String(n).padStart(2, "0"); }
    const d = new Date();
    const fname =
        `SPX_export_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(
            d.getDate()
        )}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(
            d.getSeconds()
        )}.xlsx`;

    XLSX.writeFile(wb, fname);
})();
