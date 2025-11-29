(async function(){
    // Load SheetJS
    const script = document.createElement('script');
    script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    document.head.appendChild(script);
    await new Promise(res => script.onload = res);

    function pickMainTable(){
        const tables = Array.from(document.querySelectorAll("table"));
        if(!tables.length){ alert("No tables found"); return null; }
        tables.sort((a,b)=>b.rows.length - a.rows.length);
        return tables[0];
    }

    const table = pickMainTable();
    if(!table){ return; }

    // Convert table → worksheet
    const ws = XLSX.utils.table_to_sheet(table);

    // Auto column widths
    const range = XLSX.utils.decode_range(ws['!ref']);
    const colWidths = [];

    for(let C = range.s.c; C <= range.e.c; C++){
        let maxWidth = 8;
        for(let R = range.s.r; R <= range.e.r; R++){
            const cell = ws[XLSX.utils.encode_cell({r:R, c:C})];
            if(cell && cell.v){
                const length = String(cell.v).length;
                if(length > maxWidth) maxWidth = length;
            }
        }
        colWidths.push({wch: maxWidth+2});
    }

    ws["!cols"] = colWidths;

    // Workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SPX");

    // Filename
    function pad(n){ return String(n).padStart(2,'0'); }
    const d = new Date();
    const fname =
        `SPX_export_${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}.xlsx`;

    XLSX.writeFile(wb, fname);
})();
