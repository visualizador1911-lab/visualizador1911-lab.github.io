const OUTPUT_COLUMNS = [
"ShotSpotter ID","Fecha/Hora","AÑO","MES","Tipo_Disparo","Zona","Direccion","Rondas",
"Sector","Latitud","Longitud","ESTADO","La","Lo","Lat","Long",
"TIPO_DISP_CORREGIDO","SIVVE","SGSP","TITULO_SGSP","EN_CUADRANTE","OBSERVACIONES_EVENTO",
"CÁMARAS","SE_VISUALIZA_HECHO","CASO_POSITIVO","CASO_CONFIRMADO","CASO_DE_EXITO",
"VINCULADO","DISPARO_RAFAGA","TRASLADO_VÍCT_POR_FFPP","DETENIDOS","FORMALIZADOS",
"VEHICULOS_INC","VEHICULOS_REC","ARMAS_INC","VAINAS_ENC","CALIBRE_VAINAS",
"Municiones_incautadas","CALIBRE_MUNICIONES","EN_BOCA_ DE_DROGA","INCAUTA_DROGA",
"LUGAR_DEL_EVENTO","LUGAR_DE_IMPACTO","ERROR_CONSTATADO_DIAC","TRABAJA_CIENTIFICA",
"MOTIVO_CIENTIFICA","Fecha_Ing_Planilla",
"FECHA_Y_HORA_DE_INGRESO_AL_SGSP POR_OPERADOR.","MOTIVO_DEL_INGRESO_AL_SGSP",
"Operador","Observaciones_ControlFinal"
];

const SPANISH_MONTHS = {
"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
"jul":7,"ago":8,"sep":9,"set":9,"oct":10,"nov":11,"dic":12
};

// ✅ FIX: normalize headers properly
function normalizeHeader(h) {
    return h
        ? h.toString()
            .replace(/\u00A0/g, " ")   // fix hidden Excel spaces
            .trim()
            .toLowerCase()
        : "";
}

function excelDateToJS(serial) {
    const utc_days = Math.floor(serial - 25569);
    return new Date(utc_days * 86400 * 1000);
}

function parseSpanishDate(str) {
    if (!str) return null;

    str = str.toString().toLowerCase();

    const match = str.match(/(\d{1,2})[-\/ ]([a-zñ]+)/);
    if (!match) return null;

    const day = parseInt(match[1]);
    const month = SPANISH_MONTHS[match[2]];
    if (!month) return null;

    return new Date(2026, month - 1, day);
}

function parseAnyDate(value) {
    if (!value) return null;

    if (value instanceof Date) return value;

    if (typeof value === "number") return excelDateToJS(value);

    if (typeof value === "string") {
        let d = parseSpanishDate(value);
        if (d) return d;

        let parsed = new Date(value);
        if (!isNaN(parsed)) return parsed;
    }

    return null;
}

function formatDecimal(v) {
    if (v === null || v === undefined) return "";
    return v.toString().replace(".", ",");
}

function processFile() {

    const file = document.getElementById("fileInput").files[0];
    if (!file) return alert("Selecciona un archivo");

    const reader = new FileReader();

    reader.onload = function(e) {

        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type:"array"});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const rows = XLSX.utils.sheet_to_json(sheet, {header:1});

        const headers = rows[0];

        // ✅ FIXED mapping (normalized keys)
        const inputMap = {};

        headers.forEach((h, i) => {
            const key = normalizeHeader(h);
            if (key) inputMap[key] = i;
        });

        const output = [];

        const total = rows.length - 1;

        for (let i = 1; i < rows.length; i++) {

            const row = rows[i] || [];

            let fecha = parseAnyDate(row[inputMap["fecha/hora"]]);
            let year = fecha ? fecha.getFullYear() : "";
            let month = fecha ? (fecha.getMonth() + 1) : "";

            let outRow = {};

            OUTPUT_COLUMNS.forEach(col => {

                let value = "";

                const key = normalizeHeader(col);

                if (key === "shotspotter id") {
                    value = row[inputMap["shotspotter id"]];
                }

                else if (key === "año") value = year;
                else if (key === "mes") value = month;

                else if (["la","lo","lat","long"].includes(key)) {
                    value = "";
                }

                else if (key === "latitud" || key === "longitud") {
                    let v = row[inputMap[key]];
                    value = formatDecimal(v);
                }

                else if (key === "fecha_ing_planilla") {
                    let d = parseAnyDate(row[inputMap[key]]);
                    value = d
                        ? `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`
                        : "";
                }

                else if (inputMap[key] !== undefined) {
                    value = row[inputMap[key]];
                }

                outRow[col] = value;
            });

            output.push(outRow);

            document.getElementById("progressBar").style.width =
                `${Math.floor((i / total) * 100)}%`;

            document.getElementById("status").innerText =
                `Procesando ${i}/${total}`;
        }

        const ws = XLSX.utils.json_to_sheet(output);
        const wb = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(wb, ws, "Ev_Principales");

        XLSX.writeFile(wb, "archivo_procesado.xlsx");

        document.getElementById("status").innerText =
            "✔ Proceso completado";
    };

    reader.readAsArrayBuffer(file);
}