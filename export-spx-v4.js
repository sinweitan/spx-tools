(async function () {
    console.log("SPX Tools Export Script v4 Loaded");

    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    document.head.appendChild(script);
    await new Promise((res) => (script.onload = res));

    function pickMainTable() {
        const tables = Array.from(document.querySelectorAll("table"));
        if (!tables.length) return null;
        tables.sort((a, b) => b.rows.length - a.rows.length);
        return tables[0];
    }

    const table = pickMainTable();
    if (!table) return alert("No table found.");

    const tmp = XLSX.utils.table_to_sheet(table);
    const data = XLSX.utils.sheet_to_json(tmp, { header: 1, raw: true });

    // Remove C,F,G,H,I,K
    const removeIdx = new Set([2, 5, 6, 7, 8, 10]);
    const filtered = data.map(row => row.filter((_, i) => !removeIdx.has(i)));

    // Add Code / PIN
    const finalData = filtered.map((row, r) =>
        r === 0 ? ["Code", "PIN", ...row] : ["", "", ...row]
    );

    const ws = XLSX.utils.aoa_to_sheet(finalData);

    // ------------ DATE PARSER ---------------
    function parseDate(str) {
        if (!str) return null;

        if (/\d{1,2}\/\d{1,2}\/\d{4}\s+\d/.test(str)) {
            const [d, t] = str.split(" ");
            const [dd, mm, yyyy] = d.split("/").map(Number);
            const [hh, mi] = t.split(":").map(Number);
            return new Date(yyyy, mm - 1, dd, hh, mi);
        }

        if (/\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
            const [dd, mm, yyyy] = str.split("/").map(Number);
            return new Date(yyyy, mm - 1, dd);
        }

        return null;
    }

    function convertColumn(col, withTime) {
        const r = XLSX.utils.decode_range(ws["!ref"]);
        for (let row = r.s.r + 1; row <= r.e.r; row++) {
            const addr = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = ws[addr];
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

    // ------------ BORDER + ROW HEIGHT -------------
    ws["!rows"] = new Array(finalData.length).fill(null).map(() => ({ hpx: 60 }));

    const rng = XLSX.utils.decode_range(ws["!ref"]);
    for (let R = rng.s.r; R <= rng.e.r; R++) {
        for (let C = rng.s.c; C <= rng.e.c; C++) {
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
            cell.s.alignment = { wrapText: true, vertical: "center" };
        }
    }

    // ------------ TRUE AUTO WIDTH (FINAL FIX) ------------
    ws["!cols"] = [];

    for (let C = 0; C < finalData[0].length; C++) {
        if (C < 2) {
            ws["!cols"].push({ wch: 12 });
            continue;
        }

        let max = 12;

        for (let R = 0; R < finalData.length; R++) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = ws[addr];
            if (!cell) continue;

            let text = "";

            if (cell.t === "d") {
                const dt = new Date(cell.v);
                text = cell.z.includes("h")
                    ? dt.toLocaleString("en-GB")
                    : dt.toLocaleDateString("en-GB");
            } else {
                text = String(cell.v ?? "");
            }

            max = Math.max(max, text.length + 2);
        }

        ws["!cols"].push({ wch: max });
    }

    // ------------ SAVE -------------
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SPX");

    function pad(n) { return String(n).padStart(2, "0"); }
    const d2 = new Date();
    const file =
        `SPX_export_${d2.getFullYear()}${pad(d2.getMonth() + 1)}${pad(d2.getDate())}_${pad(d2.getHours())}${pad(d2.getMinutes())}${pad(d2.getSeconds())}.xlsx`;

    XLSX.writeFile(wb, file);
})();
