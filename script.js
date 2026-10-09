/* ============================================================
   SALES CONTROL - OCT 26
   Google Sheets -> Web Dashboard
============================================================ */

"use strict";

/* ============================================================
   CONFIGURACIÓN
============================================================ */

const CSV_URL =
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-0sM4BhtW0jFQUkPZr4fkExixvpNEGe1GPisZTtFsJC-pBtpOePh0uigNq7iFAw/pub?output=csv";

const ROWS_PER_PAGE = 10;
const REFRESH_INTERVAL = 60000;
const SEARCH_DEBOUNCE = 200;
const FETCH_TIMEOUT = 20000;

/*
   true  = la proyección de comisión se calcula con las ventas filtradas
           (buscador y filtros por columna).
   false = siempre se calcula con TODAS las ventas cargadas.
*/
const PROJECTION_USES_FILTERS = true;


/* ============================================================
   ESTADO GLOBAL
============================================================ */

let headers = [];
let sales = [];
let filteredSales = [];

let currentPage = 1;
let isLoading = false;
let refreshTimer = null;
let toastTimer = null;
let searchTimer = null;
let lastFocusedElement = null;

const columnCache = new Map();


/* ============================================================
   COLUMNAS Y TABLAS DE CONFIGURACIÓN
============================================================ */

const COLUMN_ALIASES = {
    n: ["N°", "Nº", "NO", "N", "NUMERO", "NÚMERO"],
    fechaVenta: ["FECHA VENTA"],
    horaVenta: ["HORA VENTA"],
    dni: ["DNI"],
    nombre: ["NOMBRE", "NOMBRE CLIENTE"],
    tipoDespacho: ["TIPO DESPAACHO", "TIPO DESPACHO", "TIPO DE DESPACHO"],
    rango: ["RANGO"],
    rangoEntrega: ["RANGO ENTREGA"],
    departamento: ["DEPARTAMENTO"],
    provincia: ["PROVINCIA"],
    distrito: ["DISTRITO"],
    tipoOfrec: ["TIPO OFREC.", "TIPO OFREC", "TIPO OFRECIMIENTO"],
    tipoVenta: ["TIPO VENTA"],
    lineaPortar: ["LINEA A PORTAR", "LÍNEA A PORTAR"],
    tipoProducto: ["TIPO DE PROD.", "TIPO DE PROD", "TIPO DE PRODUCTO"],
    supervisor: ["SUPERVISOR"],
    estadoBO: ["ESTADO BO"],
    subestado: ["SUBESTADO"],
    ordenOT: ["ORDEN OT"],
    atencionFinal: ["ATENCIÓN FINAL", "ATENCION FINAL"],
    boValidador: ["BO VALIDADOR"],
    dniVendedor: ["DNI VENDEDOR"],
    vendedor: ["VENDEDOR"],
    estado: ["ESTADO"],
    fechaActivacion: ["FECHA ACTIVACIÓN", "FECHA ACTIVACION"],
    modalidad: ["MODALIDAD"],
    planDescuento: ["PLAN CON DESC.", "PLAN CON DESC", "PLAN CON DESCUENTO"],
    clienteEntel: ["CLIENTE ENTEL"],
    status: ["STATUS"]
};

/* Columnas sin las cuales el panel no funciona bien */
const REQUIRED_COLUMNS = ["n", "nombre", "dni", "estado"];

const MAIN_COLUMNS = [
    { key: "n", label: "N°" },
    { key: "fechaVenta", label: "FECHA VENTA" },
    { key: "dni", label: "DNI" },
    { key: "nombre", label: "NOMBRE" },
    { key: "rango", label: "RANGO" },
    { key: "distrito", label: "DISTRITO" },
    { key: "tipoOfrec", label: "TIPO OFREC." },
    { key: "lineaPortar", label: "LINEA A PORTAR" },
    { key: "estadoBO", label: "ESTADO BO" },
    { key: "ordenOT", label: "ORDEN OT", occurrence: 1 },
    { key: "estado", label: "ESTADO" },
    { key: "fechaActivacion", label: "FECHA ACTIVACIÓN" },
    { key: "modalidad", label: "MODALIDAD" }
];

const DATE_KEYS = ["fechaVenta", "fechaActivacion"];
/* Solo la columna ESTADO lleva color; las demás van como texto normal */
const STATUS_KEYS = ["estado"];

/* [etiqueta, clave, ocurrencia] */
const DETAIL_SECTIONS = [
    {
        title: "Cliente y venta",
        fields: [
            ["N°", "n"],
            ["Fecha de venta", "fechaVenta"],
            ["Hora de venta", "horaVenta"],
            ["DNI", "dni"],
            ["Nombre", "nombre"]
        ]
    },
    {
        title: "Entrega",
        fields: [
            ["Rango", "rango"],
            ["Rango entrega", "rangoEntrega"],
            ["Distrito", "distrito"],
            ["Provincia", "provincia"],
            ["Departamento", "departamento"],
            ["Tipo despacho", "tipoDespacho"]
        ]
    },
    {
        title: "Producto",
        fields: [
            ["Tipo ofrecimiento", "tipoOfrec"],
            ["Tipo de venta", "tipoVenta"],
            ["Línea a portar", "lineaPortar"],
            ["Tipo de producto", "tipoProducto"],
            ["Modalidad", "modalidad"],
            ["Plan con descuento", "planDescuento"],
            ["Cliente Entel", "clienteEntel"]
        ]
    },
    {
        title: "Gestión comercial",
        fields: [
            ["Supervisor", "supervisor"],
            ["Estado BO", "estadoBO"],
            ["Subestado", "subestado"],
            ["Orden OT", "ordenOT", 1],
            ["Atención final", "atencionFinal"],
            ["BO validador", "boValidador"]
        ]
    },
    {
        title: "Vendedor y activación",
        fields: [
            ["DNI vendedor", "dniVendedor"],
            ["Vendedor", "vendedor"],
            ["Estado", "estado"],
            ["Fecha activación", "fechaActivacion"],
            ["Status", "status"],
            ["Orden OT 2", "ordenOT", 2]
        ]
    }
];


/* ============================================================
   COMISIONES
============================================================ */

const COMMISSION_RANGES = [
    { range: 1,  pos: 5.00,  alta: 2.00,  prepago: 2.50,  multi: 8.00,  from: 20,  to: 30 },
    { range: 2,  pos: 12.50, alta: 4.00,  prepago: 6.00,  multi: 14.50, from: 31,  to: 40 },
    { range: 3,  pos: 13.50, alta: 4.50,  prepago: 6.50,  multi: 15.50, from: 41,  to: 50 },
    { range: 4,  pos: 14.50, alta: 5.00,  prepago: 7.00,  multi: 16.50, from: 51,  to: 60 },
    { range: 5,  pos: 15.50, alta: 5.50,  prepago: 7.50,  multi: 17.50, from: 61,  to: 70 },
    { range: 6,  pos: 17.00, alta: 6.00,  prepago: 8.50,  multi: 19.00, from: 71,  to: 80 },
    { range: 7,  pos: 19.00, alta: 6.00,  prepago: 10.00, multi: 22.00, from: 81,  to: 95 },
    { range: 8,  pos: 21.50, alta: 7.00,  prepago: 12.00, multi: 25.00, from: 96,  to: 110 },
    { range: 9,  pos: 24.00, alta: 8.00,  prepago: 14.00, multi: 27.00, from: 111, to: 150 },
    { range: 10, pos: 27.00, alta: 9.00,  prepago: 16.00, multi: 30.00, from: 151, to: 199 },
    { range: 11, pos: 30.00, alta: 12.00, prepago: 18.00, multi: 33.00, from: 200, to: Infinity }
];

const TOP_RANGE = COMMISSION_RANGES[COMMISSION_RANGES.length - 1];
const TARGET_SALES = TOP_RANGE.from;

/*
   AJUSTES DE LA COMISIÓN (castigos y bonificadores)
   - El % se aplica sobre la comisión base (POS MONO + ALTA + PREPAGO + MULTI).
   - Se evalúan de arriba hacia abajo; gana la primera regla que cumple.
*/

/* Tableau: alcance = órdenes OT activas / total de órdenes OT */
const TABLEAU_RULES = [
    { min: 80,        pct: 15,  label: "≥ 80%" },
    { min: 75,        pct: 5,   label: "≥ 75%" },
    { min: 70,        pct: -5,  label: "< 75%" },
    { min: -Infinity, pct: -15, label: "< 70%" }
];

/* Multipedido: alcance = ventas multipedido activas / ventas activas */
const MULTI_RULES = [
    { min: 35,        pct: 0,   label: "≥ 35%" },
    { min: 30,        pct: -5,  label: "< 35%" },
    { min: -Infinity, pct: -10, label: "< 30%" }
];

/* Redondea a 2 decimales para que 30% sea 30% y no 29.999999% */
function round2(value) {
    return Math.round(value * 100) / 100;
}

function findRule(rules, value) {
    return rules.find(rule => round2(value) >= rule.min);
}


/* ============================================================
   UTILIDADES
============================================================ */

const $ = id => document.getElementById(id);

function normalize(value) {
    return String(value ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase()
        .trim();
}

function normalizeHeader(value) {
    return normalize(value)
        .replace(/[°º]/g, "")
        .replace(/[.:;,()[\]{}]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function cleanValue(value) {
    return String(value ?? "").replace(/\r/g, "").trim();
}

function escapeHTML(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function setText(id, value) {
    const element = $(id);

    if (element) {
        element.textContent = value;
    }
}

function clamp(value, min = 0, max = 100) {
    return Math.max(min, Math.min(max, Number(value) || 0));
}

function setWidth(id, percentageValue) {
    const element = $(id);

    if (element) {
        element.style.width = `${clamp(percentageValue)}%`;
    }
}

function percentage(value, total) {
    return total ? (value / total) * 100 : 0;
}

function formatPercentage(value) {
    if (!Number.isFinite(value)) {
        return "0%";
    }

    return `${value
        .toFixed(2)
        .replace(/\.00$/, "")
        .replace(/(\.\d)0$/, "$1")}%`;
}

const moneyFormatter = new Intl.NumberFormat("es-PE", {
    style: "currency",
    currency: "PEN",
    minimumFractionDigits: 2
});

function formatMoney(value) {
    return moneyFormatter.format(value || 0);
}

function formatSignedMoney(value) {
    if (value > 0) return `+${formatMoney(value)}`;
    if (value < 0) return `−${formatMoney(Math.abs(value))}`;

    return formatMoney(0);
}

function formatDateValue(value) {
    const text = cleanValue(value);
    const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);

    return match
        ? `${match[3]}/${match[2]}/${match[1]}`
        : text;
}


/* ============================================================
   CSV
============================================================ */

function detectDelimiter(text) {
    const firstLine = text.split(/\r?\n/, 1)[0] || "";

    let comma = 0;
    let semicolon = 0;
    let insideQuotes = false;

    for (const char of firstLine) {
        if (char === '"') {
            insideQuotes = !insideQuotes;
        } else if (!insideQuotes) {
            if (char === ",") comma++;
            if (char === ";") semicolon++;
        }
    }

    return semicolon > comma ? ";" : ",";
}

function parseCSV(text, delimiter) {
    const rows = [];

    let row = [];
    let value = "";
    let insideQuotes = false;

    const pushRow = () => {
        row.push(value);

        if (row.some(cell => cleanValue(cell) !== "")) {
            rows.push(row);
        }

        row = [];
        value = "";
    };

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const next = text[i + 1];

        if (char === '"') {
            if (insideQuotes && next === '"') {
                value += '"';
                i++;
            } else {
                insideQuotes = !insideQuotes;
            }
        } else if (char === delimiter && !insideQuotes) {
            row.push(value);
            value = "";
        } else if ((char === "\n" || char === "\r") && !insideQuotes) {
            if (char === "\r" && next === "\n") {
                i++;
            }

            pushRow();
        } else {
            value += char;
        }
    }

    if (value !== "" || row.length > 0) {
        pushRow();
    }

    return rows;
}


/* ============================================================
   ACCESO A COLUMNAS (con caché)
============================================================ */

function computeColumnIndex(aliases, occurrence) {
    const wanted = aliases.map(normalizeHeader);
    const normalizedHeaders = headers.map(normalizeHeader);

    const exact = [];
    const partial = [];

    normalizedHeaders.forEach((header, index) => {
        if (!header) {
            return;
        }

        if (wanted.includes(header)) {
            exact.push(index);
        }

        /* Coincidencia parcial solo para textos largos: evita falsos positivos */
        const isPartial = wanted.some(alias =>
            alias.length >= 4 &&
            header.length >= 4 &&
            (header.includes(alias) || alias.includes(header))
        );

        if (isPartial) {
            partial.push(index);
        }
    });

    /* "last": la última columna con ese nombre (ej. el ORDEN OT del final) */
    if (occurrence === "last") {
        const list = exact.length ? exact : partial;

        return list.length ? list[list.length - 1] : -1;
    }

    if (exact.length >= occurrence) {
        return exact[occurrence - 1];
    }

    if (partial.length >= occurrence) {
        return partial[occurrence - 1];
    }

    return -1;
}

function findColumnIndex(aliases, occurrence = 1) {
    if (!headers.length) {
        return -1;
    }

    const cacheKey = `${aliases.join("|")}#${occurrence}`;

    if (!columnCache.has(cacheKey)) {
        columnCache.set(cacheKey, computeColumnIndex(aliases, occurrence));
    }

    return columnCache.get(cacheKey);
}

function resolveAliases(keyOrAliases) {
    if (Array.isArray(keyOrAliases)) {
        return keyOrAliases;
    }

    return COLUMN_ALIASES[keyOrAliases] || [keyOrAliases];
}

function getValue(row, keyOrAliases, occurrence = 1) {
    const index = findColumnIndex(resolveAliases(keyOrAliases), occurrence);

    return index < 0 || index >= row.length
        ? ""
        : cleanValue(row[index]);
}

function missingColumns() {
    return REQUIRED_COLUMNS.filter(
        key => findColumnIndex(resolveAliases(key)) < 0
    );
}


/* ============================================================
   CLASIFICACIÓN DE VENTAS
============================================================ */

function isRealSale(row) {
    const hasNumber = getValue(row, "n").replace(/\D/g, "") !== "";

    if (!hasNumber) {
        return false;
    }

    const hasName = normalize(getValue(row, "nombre")).length > 0;
    const hasDni = getValue(row, "dni").replace(/\D/g, "").length >= 6;

    return hasName || hasDni;
}

/* Clasifica un texto de estado. Devuelve null si no reconoce nada. */
function classifyStatusText(text) {
    /* \b evita que "INACTIVO" o "INVALIDO" cuenten como activas */
    if (
        /\b(ACTIV|VALID)/.test(text) &&
        !/\bNO (ACTIV|VALID)/.test(text)
    ) {
        return "active";
    }

    if (/PROGRES|PENDIENT|EN PROCESO|EN CURSO/.test(text)) {
        return "progress";
    }

    return null;
}

/*
   Devuelve: "cancelled" | "noRecoge" | "active" | "progress"

   - Cancelada / No recoge: se detecta en cualquiera de las columnas
     de estado (como antes).
   - Activa / En progreso: manda la columna ESTADO. Así una venta con
     ESTADO "EN PROGRESO" no pasa a activa solo porque ESTADO BO
     diga "VALIDA".
*/
function getOperationalStatus(row) {
    const estado = normalize(getValue(row, "estado"));

    const combined = [
        estado,
        ...["estadoBO", "subestado", "atencionFinal", "status"]
            .map(key => normalize(getValue(row, key)))
    ].join(" | ");

    if (/\b(CANCEL|RECHAZ)/.test(combined)) {
        return "cancelled";
    }

    if (/NO ?RECOGE/.test(combined)) {
        return "noRecoge";
    }

    return (
        classifyStatusText(estado) ||
        classifyStatusText(combined) ||
        "progress"
    );
}

/*
   Cantidad de órdenes OT de una venta, leída de la ÚLTIMA columna "ORDEN OT".
   - Número chico (1, 2, 3...): es la cantidad de órdenes.
   - Uno o varios números largos (ids de orden): cuenta cada id.
   - Vacío: 0 órdenes.
*/
function getOtCount(row) {
    const raw = getValue(row, "ordenOT", "last");

    if (!raw) {
        return 0;
    }

    const small = raw.match(/^\s*(\d{1,3})(?!\d)/);

    if (small) {
        return parseInt(small[1], 10);
    }

    const ids = raw.match(/\d{6,}/g);

    return ids ? ids.length : 0;
}

function getChannel(row) {
    const combined = normalize(
        `${getValue(row, "tipoDespacho")} ${getValue(row, "modalidad")}`
    );

    if (combined.includes("TIENDA") || combined.includes("STORE")) {
        return "tienda";
    }

    if (combined.includes("DELIVERY") || combined.includes("DOMICILIO")) {
        return "delivery";
    }

    return "other";
}

function getProductText(row) {
    return normalize(
        [
            getValue(row, "tipoVenta"),
            getValue(row, "tipoProducto"),
            getValue(row, "tipoOfrec"),
            getValue(row, "modalidad"),
            getValue(row, "lineaPortar")
        ].join(" ")
    );
}

/*
   Multipedido = la columna TIPO OFREC. dice MULTILINEA (X2, X3, X4, X5...).
   Una venta regular solo dice "REGULAR" y nunca cuenta como multipedido.
*/
function isMultipedido(row) {
    return /MULTI\s*LINEA/.test(normalize(getValue(row, "tipoOfrec")));
}

function getCommissionType(productText, isMulti) {
    if (isMulti) {
        return "multi";
    }

    if (productText.includes("PREPAGO")) {
        return "prepago";
    }

    if (/\bALTA\b|LLAA/.test(productText)) {
        return "alta";
    }

    return "pos";
}

/*
   Cada venta se clasifica UNA sola vez al cargar los datos.
   Antes se recalculaba en cada render (miles de veces).
*/
function buildSale(row) {
    const productText = getProductText(row);
    const isMulti = isMultipedido(row);

    const searchText = [
        "n", "fechaVenta", "horaVenta", "dni", "nombre", "distrito",
        "tipoOfrec", "lineaPortar", "estado", "vendedor", "supervisor"
    ]
        .map(key => getValue(row, key))
        .concat(getValue(row, "ordenOT", 1))
        .map(normalize)
        .join(" ");

    /* Valores de las columnas que se pueden filtrar */
    const fields = {};

    FILTERS.forEach(filter => {
        const raw = getValue(row, filter.key, filter.occurrence || 1);

        fields[filter.key] = filter.type === "date"
            ? toISODate(raw)
            : raw;
    });

    return {
        row,
        status: getOperationalStatus(row),
        channel: getChannel(row),
        isMulti,
        commissionType: getCommissionType(productText, isMulti),
        otCount: getOtCount(row),
        fields,
        searchText
    };
}

/* Resumen de estados. "cancelled" del panel = canceladas + no recoge. */
function summarize(list) {
    const result = {
        total: list.length,
        active: 0,
        progress: 0,
        cancelledOnly: 0,
        noRecoge: 0
    };

    for (const sale of list) {
        if (sale.status === "active") result.active++;
        else if (sale.status === "progress") result.progress++;
        else if (sale.status === "cancelled") result.cancelledOnly++;
        else if (sale.status === "noRecoge") result.noRecoge++;
    }

    result.cancelled = result.cancelledOnly + result.noRecoge;

    return result;
}


/* ============================================================
   CARGA DE DATOS
============================================================ */

function buildCsvUrl() {
    const url = new URL(CSV_URL);
    url.searchParams.set("t", Date.now());

    return url.toString();
}

async function cargarDatos(showToastMessage = false) {
    if (isLoading) {
        return;
    }

    isLoading = true;

    const refreshButton = $("refreshButton");

    if (refreshButton) {
        refreshButton.disabled = true;
    }

    setConnection("loading", "Actualizando...");

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), FETCH_TIMEOUT);

    try {
        const response = await fetch(buildCsvUrl(), {
            cache: "no-store",
            signal: controller.signal
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const text = (await response.text()).replace(/^\uFEFF/, "");

        if (text.trim().length < 10) {
            throw new Error("Google Sheets no devolvió datos.");
        }

        const parsed = parseCSV(text, detectDelimiter(text));

        if (parsed.length < 2) {
            throw new Error("No se encontraron filas.");
        }

        headers = parsed[0].map(cleanValue);
        columnCache.clear();

        const missing = missingColumns();

        if (missing.length) {
            console.warn("Columnas no encontradas en la hoja:", missing);
        }

        const otColumns = headers
            .map(normalizeHeader)
            .filter(header => header === "ORDEN OT").length;

        if (otColumns < 2) {
            console.warn(
                "Solo hay una columna ORDEN OT: Tableau contará 1 orden por venta con número de orden."
            );
        }

        sales = parsed
            .slice(1)
            .filter(isRealSale)
            .map(buildSale);

        const totalOt = sales.reduce((sum, sale) => sum + sale.otCount, 0);

        cargarFiltros();
        aplicarFiltros();

        setConnection("connected", "Conectado");

        setText(
            "lastUpdate",
            `Actualizado ${new Date().toLocaleString("es-PE", {
                dateStyle: "short",
                timeStyle: "short"
            })}`
        );

        if (missing.length) {
            mostrarToast(
                `Faltan columnas en la hoja: ${missing.join(", ")}.`
            );
        } else if (sales.length && totalOt === 0) {
            mostrarToast(
                "No se leyeron órdenes en la columna ORDEN OT. Revisa la hoja."
            );
        } else if (showToastMessage) {
            mostrarToast(`Datos actualizados: ${sales.length} ventas.`);
        }

    } catch (error) {
        console.error("Error cargando Google Sheets:", error);

        const reason = error.name === "AbortError"
            ? "La conexión tardó demasiado."
            : "Revisa tu conexión e inténtalo de nuevo.";

        setConnection("error", "Error de conexión");

        if (!sales.length) {
            setTableMessage("No se pudieron cargar las ventas", reason);
            $("tableEmpty").hidden = false;
        }

        mostrarToast(`No se pudieron actualizar los datos. ${reason}`);

    } finally {
        clearTimeout(timeout);
        isLoading = false;

        if (refreshButton) {
            refreshButton.disabled = false;
        }
    }
}

function setConnection(type, text) {
    setText("connectionStatus", text);

    const dot = $("connectionDot");

    if (dot) {
        dot.className = type === "connected" || type === "error" || type === "loading"
            ? type
            : "";
    }
}


/* ============================================================
   FILTROS
============================================================ */

/* type: "select" (lista de valores), "text" (contiene) o "date" (desde/hasta) */
const FILTERS = [
    { key: "fechaVenta", label: "Fecha de venta", type: "date" },
    { key: "dni", label: "DNI", type: "text" },
    { key: "rango", label: "Rango", type: "select" },
    { key: "distrito", label: "Distrito", type: "select" },
    { key: "tipoOfrec", label: "Tipo ofrecimiento", type: "select" },
    { key: "lineaPortar", label: "Línea a portar", type: "select" },
    { key: "estadoBO", label: "Estado BO", type: "select" },
    { key: "ordenOT", label: "Orden OT", type: "text", occurrence: 1 },
    { key: "estado", label: "Estado", type: "select" },
    { key: "fechaActivacion", label: "Fecha de activación", type: "date" },
    { key: "modalidad", label: "Modalidad", type: "select" }
];

const EMPTY_OPTION = "__EMPTY__";

const filterId = (filter, suffix = "") => `filter-${filter.key}${suffix}`;

/* Convierte 2026-10-05 o 05/10/2026 a "2026-10-05" (comparable) */
function toISODate(value) {
    const text = cleanValue(value);
    const pad = number => String(number).padStart(2, "0");

    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);

    if (match) {
        return `${match[1]}-${pad(match[2])}-${pad(match[3])}`;
    }

    match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);

    if (match) {
        const year = match[3].length === 2 ? `20${match[3]}` : match[3];

        return `${year}-${pad(match[2])}-${pad(match[1])}`;
    }

    return "";
}

function createFilterField(filter) {
    const id = filterId(filter);
    const label = escapeHTML(filter.label);
    const field = document.createElement("div");

    field.className = filter.type === "date" ? "field field-date" : "field";

    if (filter.type === "date") {
        field.innerHTML = `
            <span class="field-label" id="${id}-label">${label}</span>
            <div class="date-range" role="group" aria-labelledby="${id}-label">
                <label>Desde <input type="date" id="${id}-from"></label>
                <label>Hasta <input type="date" id="${id}-to"></label>
            </div>
        `;
    } else if (filter.type === "text") {
        field.innerHTML = `
            <label for="${id}">${label}</label>
            <input type="text" id="${id}" autocomplete="off" placeholder="Contiene...">
        `;
    } else {
        field.innerHTML = `
            <label for="${id}">${label}</label>
            <select id="${id}"><option value="ALL">Todos</option></select>
        `;
    }

    return field;
}

function construirFiltros() {
    const container = $("advancedFilters");

    if (container) {
        container.replaceChildren(...FILTERS.map(createFilterField));
    }
}

/* Llena cada lista con los valores reales de la hoja y conserva la selección */
function cargarFiltros() {
    FILTERS.filter(filter => filter.type === "select").forEach(filter => {
        const select = $(filterId(filter));

        if (!select) {
            return;
        }

        const current = select.value;
        const values = new Set();

        let hasEmpty = false;

        sales.forEach(sale => {
            const value = sale.fields[filter.key];

            if (value) {
                values.add(value);
            } else {
                hasEmpty = true;
            }
        });

        select.replaceChildren(new Option("Todos", "ALL"));

        [...values]
            .sort((a, b) => a.localeCompare(b, "es", { numeric: true }))
            .forEach(value => select.add(new Option(value, value)));

        if (hasEmpty) {
            select.add(new Option("(Sin dato)", EMPTY_OPTION));
        }

        /* Conserva la selección aunque el valor ya no exista en los datos */
        [current, appliedValues[select.id]].forEach(value => {
            if (
                value &&
                value !== "ALL" &&
                ![...select.options].some(option => option.value === value)
            ) {
                select.add(new Option(value, value));
            }
        });

        select.value = current || "ALL";
    });
}

/* Valores aplicados de los filtros por columna: id del control -> valor */
let appliedValues = {};

function leerControles() {
    const values = {};

    $("advancedFilters")
        ?.querySelectorAll("input, select")
        .forEach(control => {
            values[control.id] = control.value;
        });

    return values;
}

function escribirControles(values) {
    $("advancedFilters")
        ?.querySelectorAll("input, select")
        .forEach(control => {
            const isSelect = control.tagName === "SELECT";
            const value = values[control.id] ?? (isSelect ? "ALL" : "");

            control.value = value;

            if (isSelect && control.value !== value) {
                control.value = "ALL";
            }
        });
}

/* Devuelve solo los filtros aplicados (no los que se están editando) */
function leerFiltros() {
    const active = [];

    FILTERS.forEach(filter => {
        const id = filterId(filter);

        if (filter.type === "date") {
            const from = appliedValues[`${id}-from`] || "";
            const to = appliedValues[`${id}-to`] || "";

            if (from || to) {
                const parts = [];

                if (from) parts.push(`desde ${formatDateValue(from)}`);
                if (to) parts.push(`hasta ${formatDateValue(to)}`);

                active.push({
                    filter,
                    from,
                    to,
                    ids: [`${id}-from`, `${id}-to`],
                    display: parts.join(" ")
                });
            }
        } else if (filter.type === "text") {
            const raw = cleanValue(appliedValues[id]);

            if (raw) {
                active.push({
                    filter,
                    text: normalize(raw),
                    ids: [id],
                    display: raw
                });
            }
        } else {
            const value = appliedValues[id] || "ALL";

            if (value !== "ALL") {
                active.push({
                    filter,
                    value,
                    ids: [id],
                    display: value === EMPTY_OPTION ? "Sin dato" : value
                });
            }
        }
    });

    return active;
}

function matchesFilter(sale, active) {
    const { filter } = active;
    const value = sale.fields[filter.key];

    if (filter.type === "date") {
        return Boolean(value) &&
            (!active.from || value >= active.from) &&
            (!active.to || value <= active.to);
    }

    if (filter.type === "text") {
        return normalize(value).includes(active.text);
    }

    return active.value === EMPTY_OPTION
        ? !value
        : value === active.value;
}

/* Etiquetas con los filtros activos; cada una se puede quitar con la × */
function renderFiltrosActivos(active) {
    const container = $("activeFilters");

    if (!container) {
        return;
    }

    container.replaceChildren(
        ...active.map(item => {
            const chip = document.createElement("span");
            const text = document.createElement("span");
            const remove = document.createElement("button");

            chip.className = "filter-chip";
            text.textContent = `${item.filter.label}: ${item.display}`;

            remove.type = "button";
            remove.className = "filter-chip-remove";
            remove.textContent = "×";
            remove.setAttribute("aria-label", `Quitar filtro ${item.filter.label}`);
            remove.addEventListener("click", () => limpiarFiltro(item.ids));

            chip.append(text, remove);

            return chip;
        })
    );
}

function limpiarFiltro(ids) {
    ids.forEach(id => {
        appliedValues[id] = $(id)?.tagName === "SELECT" ? "ALL" : "";
    });

    escribirControles(appliedValues);
    resetPageAndFilter();
}

function aplicarFiltros() {
    const search = normalize($("searchInput")?.value);
    const active = leerFiltros();

    filteredSales = sales.filter(sale =>
        (!search || sale.searchText.includes(search)) &&
        active.every(filter => matchesFilter(sale, filter))
    );

    const badge = $("filtersBadge");

    if (badge) {
        badge.textContent = active.length;
        badge.hidden = active.length === 0;
    }

    const count = active.length + (search ? 1 : 0);

    setText(
        "clearFilters",
        count ? `Limpiar filtros (${count})` : "Limpiar filtros"
    );

    renderFiltrosActivos(active);
    renderAll();
}


/* ---------- Panel "Filtros" ---------- */

function abrirFiltros() {
    const panel = $("filtersPanel");

    if (!panel) {
        return;
    }

    escribirControles(appliedValues);

    panel.hidden = false;
    $("filtersToggle")?.setAttribute("aria-expanded", "true");
}

/* Cierra el panel y descarta lo que no se haya aplicado */
function cerrarFiltros() {
    const panel = $("filtersPanel");

    if (!panel || panel.hidden) {
        return;
    }

    panel.hidden = true;
    $("filtersToggle")?.setAttribute("aria-expanded", "false");

    escribirControles(appliedValues);
}

function aplicarPanel() {
    appliedValues = leerControles();

    resetPageAndFilter();
    cerrarFiltros();
}

function limpiarTodosPanel() {
    appliedValues = {};

    escribirControles({});
    resetPageAndFilter();
}

function resetPageAndFilter() {
    currentPage = 1;
    aplicarFiltros();
}


/* ============================================================
   RENDER GENERAL
============================================================ */

function renderAll() {
    actualizarKPIs();
    renderTabla();
    renderAnalytics();
    calculateChannelKPIs(filteredSales);
    calculateProjection(PROJECTION_USES_FILTERS ? filteredSales : sales);
}

function actualizarKPIs() {
    const stats = summarize(filteredSales);

    setText("kpiTotal", stats.total);
    setText("kpiActive", stats.active);
    setText("kpiProgress", stats.progress);
    setText("kpiCancelled", stats.cancelled);
}


/* ============================================================
   TABLA
============================================================ */

function setTableMessage(title, text) {
    setText("tableEmptyTitle", title);
    setText("tableEmptyText", text);
}

function renderTableHead() {
    const head = $("tableHead");

    const cells = MAIN_COLUMNS
        .map(column => `<th scope="col">${escapeHTML(column.label)}</th>`)
        .join("");

    head.innerHTML = `<tr>${cells}<th scope="col">ACCIÓN</th></tr>`;
}

function renderTabla() {
    const body = $("tableBody");
    const empty = $("tableEmpty");

    if (!body || !$("tableHead")) {
        return;
    }

    renderTableHead();

    const totalPages = Math.max(
        1,
        Math.ceil(filteredSales.length / ROWS_PER_PAGE)
    );

    currentPage = clamp(currentPage, 1, totalPages);

    const start = (currentPage - 1) * ROWS_PER_PAGE;
    const pageRows = filteredSales.slice(start, start + ROWS_PER_PAGE);

    setText(
        "paginationInfo",
        `${filteredSales.length} ${filteredSales.length === 1 ? "registro" : "registros"}`
    );

    body.replaceChildren();

    if (!pageRows.length) {
        setTableMessage(
            "No se encontraron ventas",
            "Prueba cambiando los filtros o la búsqueda."
        );

        empty.hidden = false;
        actualizarPaginacion(0, 0);

        return;
    }

    empty.hidden = true;

    const fragment = document.createDocumentFragment();

    pageRows.forEach(sale => fragment.appendChild(createTableRow(sale)));

    body.appendChild(fragment);

    actualizarPaginacion(currentPage, totalPages);
}

function createTableRow(sale) {
    const tr = document.createElement("tr");

    MAIN_COLUMNS.forEach(column => {
        const td = document.createElement("td");

        let value = getValue(sale.row, column.key, column.occurrence || 1);

        if (DATE_KEYS.includes(column.key)) {
            value = formatDateValue(value);
        }

        if (STATUS_KEYS.includes(column.key)) {
            td.innerHTML = renderStatus(value);
        } else if (value) {
            td.textContent = value;
        } else {
            td.innerHTML = `<span class="table-empty-value">—</span>`;
        }

        if (column.key === "n") {
            td.classList.add("table-number");
        }

        tr.appendChild(td);
    });

    const actionTd = document.createElement("td");
    const button = document.createElement("button");

    button.type = "button";
    button.className = "action-button";
    button.textContent = "Ver detalle";
    button.addEventListener("click", () => mostrarDetalle(sale));

    actionTd.appendChild(button);
    tr.appendChild(actionTd);

    return tr;
}

function renderStatus(value) {
    const text = cleanValue(value);

    if (!text) {
        return `<span class="status-badge status-neutral">—</span>`;
    }

    const normalized = normalize(text);

    let className = "status-neutral";

    if (/\b(CANCEL|RECHAZ)/.test(normalized)) {
        className = "status-cancelled";
    } else if (/\b(ACTIV|VALID)/.test(normalized)) {
        className = "status-active";
    } else if (/PROGRES|PENDIENT/.test(normalized)) {
        className = "status-progress";
    }

    return `<span class="status-badge ${className}">${escapeHTML(text)}</span>`;
}


/* ============================================================
   PAGINACIÓN
============================================================ */

function createPageButton(text, disabled, label, action) {
    const button = document.createElement("button");

    button.type = "button";
    button.className = "page-button";
    button.textContent = text;
    button.disabled = disabled;
    button.setAttribute("aria-label", label);
    button.addEventListener("click", action);

    return button;
}

function goToPage(page) {
    currentPage = page;
    renderTabla();

    $("tableBody")
        ?.closest(".section-card")
        ?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function actualizarPaginacion(page, totalPages) {
    const pagination = $("pagination");

    if (!pagination) {
        return;
    }

    pagination.replaceChildren();

    if (totalPages <= 1) {
        return;
    }

    const maxButtons = 7;

    let start = Math.max(1, page - Math.floor(maxButtons / 2));
    const end = Math.min(totalPages, start + maxButtons - 1);

    start = Math.max(1, end - maxButtons + 1);

    pagination.appendChild(
        createPageButton("‹", page <= 1, "Página anterior", () => goToPage(page - 1))
    );

    for (let i = start; i <= end; i++) {
        const button = createPageButton(i, false, `Página ${i}`, () => goToPage(i));

        if (i === page) {
            button.classList.add("active");
            button.setAttribute("aria-current", "page");
        }

        pagination.appendChild(button);
    }

    pagination.appendChild(
        createPageButton("›", page >= totalPages, "Página siguiente", () => goToPage(page + 1))
    );
}


/* ============================================================
   ANALYTICS
============================================================ */

function renderAnalytics() {
    const stats = summarize(filteredSales);

    const rows = [
        ["Active", stats.active],
        ["Progress", stats.progress],
        ["Cancelled", stats.cancelled]
    ];

    setText("analyticsTotal", stats.total);

    rows.forEach(([name, count]) => {
        const pct = percentage(count, stats.total);

        setText(`analytics${name}Pct`, formatPercentage(pct));
        setWidth(`analyticsBar${name}`, pct);
    });
}


/* ============================================================
   KPIs POR CANAL
============================================================ */

function setPie(id, percentageValue, centerValue, label) {
    const element = $(id);

    if (!element) {
        return;
    }

    element.style.setProperty("--value", clamp(percentageValue));

    const value = element.querySelector(".pie-value");
    const labelElement = element.querySelector(".pie-label");

    if (value) value.textContent = centerValue;
    if (labelElement) labelElement.textContent = label;
}

/* Escribe "<prefijo>Q" y "<prefijo>Pct" */
function setLine(prefix, count, total) {
    setText(`${prefix}Q`, count);
    setText(`${prefix}Pct`, formatPercentage(percentage(count, total)));
}

function setChannelPies(chartId, totalChartId, stats) {
    const activePct = percentage(stats.active, stats.total);

    setPie(chartId, activePct, formatPercentage(activePct), "ACTIVAS");
    setPie(totalChartId, stats.total > 0 ? 100 : 0, stats.total, "TOTAL");
}

function calculateChannelKPIs(currentSales) {
    renderTiendaKPI(currentSales.filter(sale => sale.channel === "tienda"));
    renderMultipedidoKPI(currentSales.filter(sale => sale.isMulti));
    renderDeliveryKPI(currentSales.filter(sale => sale.channel === "delivery"));
    renderTableauKPI(currentSales);
}

/* Tableau: el 100% es la suma de órdenes OT (no el número de ventas) */
function getTableauStats(list) {
    const stats = { total: 0, active: 0, cancelled: 0, progress: 0 };

    for (const sale of list) {
        const orders = sale.otCount;

        if (!orders) {
            continue;
        }

        stats.total += orders;

        if (sale.status === "active") stats.active += orders;
        else if (sale.status === "progress") stats.progress += orders;
        else stats.cancelled += orders; /* cancelada + no recoge */
    }

    stats.alcance = percentage(stats.active, stats.total);

    return stats;
}

function renderTableauKPI(list) {
    const stats = getTableauStats(list);

    setLine("tableauActivas", stats.active, stats.total);
    setLine("tableauCanceladas", stats.cancelled, stats.total);
    setLine("tableauProgreso", stats.progress, stats.total);

    setText("tableauTotalQ", stats.total);
    setText("tableauTotalPct", stats.total ? "100%" : "0%");

    setPie("chartTableau", stats.alcance, formatPercentage(stats.alcance), "ACTIVAS");
    setPie("chartTableauTotal", stats.total > 0 ? 100 : 0, stats.total, "TOTAL OT");
}

function renderTiendaKPI(rows) {
    const stats = summarize(rows);

    setLine("tiendaActivas", stats.active, stats.total);
    /* Canceladas = canceladas + "no recoge", igual que en el resto del panel */
    setLine("tiendaNoRecoge", stats.cancelled, stats.total);
    setLine("tiendaProgreso", stats.progress, stats.total);

    setText("tiendaEfectividadQ", stats.total);
    setText(
        "tiendaEfectividadPct",
        formatPercentage(percentage(stats.active, stats.total))
    );

    setChannelPies("chartTienda", "chartTiendaTotal", stats);
}

function renderMultipedidoKPI(rows) {
    const stats = summarize(rows);

    setLine("multiActivas", stats.active, stats.total);
    setLine("multiCanceladas", stats.cancelled, stats.total);
    setLine("multiProgreso", stats.progress, stats.total);

    setText("multiTotalQ", stats.total);
    setText("multiTotalPct", stats.total ? "100%" : "0%");

    setChannelPies("chartMulti", "chartMultiTotal", stats);
}

function renderDeliveryKPI(rows) {
    const stats = summarize(rows);

    setLine("deliveryActivas", stats.active, stats.total);
    setLine("deliveryCanceladas", stats.cancelled, stats.total);
    setLine("deliveryProgreso", stats.progress, stats.total);

    setText("deliveryTotalQ", stats.total);
    setText(
        "deliveryEfectividadPct",
        formatPercentage(percentage(stats.active, stats.total))
    );

    setChannelPies("chartDelivery", "chartDeliveryTotal", stats);
}


/* ============================================================
   PROYECCIÓN DE COMISIONES
============================================================ */

function getCommissionRange(total) {
    return COMMISSION_RANGES.find(
        range => total >= range.from && total <= range.to
    ) || null;
}

function getNextCommissionRange(total) {
    return COMMISSION_RANGES.find(range => total < range.from) || null;
}

function calculateProjection(currentSales) {
    const activeSales = currentSales.filter(sale => sale.status === "active");
    const activeCount = activeSales.length;

    const currentRange = getCommissionRange(activeCount);

    /* Con menos de 20 ventas se usa el Rango 1 solo como referencia */
    const pricingRange = currentRange || COMMISSION_RANGES[0];

    const totals = { pos: 0, alta: 0, prepago: 0, multi: 0 };
    const counts = { pos: 0, alta: 0, prepago: 0, multi: 0 };

    activeSales.forEach(sale => {
        totals[sale.commissionType] += pricingRange[sale.commissionType];
        counts[sale.commissionType]++;
    });

    const baseCommission =
        totals.pos + totals.alta + totals.prepago + totals.multi;

    /* Ajuste por Tableau (sobre órdenes OT) */
    const tableauStats = getTableauStats(currentSales);
    const tableauRule = tableauStats.total
        ? findRule(TABLEAU_RULES, tableauStats.alcance)
        : null;

    /* Ajuste por Multipedido (sobre ventas activas) */
    const multiActive = activeSales.filter(sale => sale.isMulti).length;
    const multiAlcance = percentage(multiActive, activeCount);
    const multiRule = activeCount
        ? findRule(MULTI_RULES, multiAlcance)
        : null;

    const tableauAmount = baseCommission * ((tableauRule?.pct || 0) / 100);
    const multiAmount = baseCommission * ((multiRule?.pct || 0) / 100);

    const totalCommission = baseCommission + tableauAmount + multiAmount;

    const missing = Math.max(TARGET_SALES - activeCount, 0);
    const progressPct = Math.min(100, percentage(activeCount, TARGET_SALES));

    setText("projectionActiveSales", activeCount);
    setText(
        "projectionCurrentRange",
        currentRange ? `Rango ${currentRange.range}` : "Aún no alcanzado"
    );
    setText("projectionMissingLabel", `Faltan para Rango ${TOP_RANGE.range}`);
    setText("projectionMissingRange", missing);

    setText("projectionTotal", formatMoney(totalCommission));
    setText("projectionPos", formatMoney(totals.pos));
    setText("projectionAlta", formatMoney(totals.alta));
    setText("projectionPrepago", formatMoney(totals.prepago));
    setText("projectionMulti", formatMoney(totals.multi));

    /* Cuántas ventas de cada tipo y cuánto vale cada una según el rango */
    const rangeLabel = currentRange
        ? `Tarifa Rango ${currentRange.range}`
        : `Tarifa Rango ${pricingRange.range} (referencia)`;

    [
        ["projectionPos", "pos"],
        ["projectionAlta", "alta"],
        ["projectionPrepago", "prepago"],
        ["projectionMulti", "multi"]
    ].forEach(([id, type]) => {
        setText(`${id}Count`, `${counts[type]} ${counts[type] === 1 ? "venta" : "ventas"}`);
        setText(`${id}Price`, `${formatMoney(pricingRange[type])} c/u`);
        setText(`${id}Note`, rangeLabel);
    });

    renderAdjustment(
        "projectionTableau",
        tableauAmount,
        tableauRule,
        tableauStats.alcance,
        `${tableauStats.active} de ${tableauStats.total} órdenes OT activas`,
        "Sin órdenes OT en la hoja"
    );

    renderAdjustment(
        "projectionMultipedido",
        multiAmount,
        multiRule,
        multiAlcance,
        `${multiActive} multipedido activas de ${activeCount} activas`,
        "Sin ventas activas"
    );

    setText(
        "projectionNote",
        currentRange
            ? `Tarifa aplicada: Rango ${currentRange.range}.`
            : `Aún no se alcanzan ${COMMISSION_RANGES[0].from} ventas. Se usa Rango 1 como referencia de proyección.`
    );

    setText("rangeProgressTitle", `Progreso hacia Rango ${TOP_RANGE.range}`);
    setText("rangeProgressText", `${activeCount} de ${TARGET_SALES} ventas`);
    setText("rangeProgressCurrent", `${activeCount} / ${TARGET_SALES}`);
    setWidth("rangeProgressBar", progressPct);

    const track = $("rangeProgressTrack");

    if (track) {
        track.setAttribute("aria-valuemax", TARGET_SALES);
        track.setAttribute("aria-valuenow", Math.min(activeCount, TARGET_SALES));
    }

    renderCommissionTable(currentRange, activeCount);
}

/* Caja de ajuste: monto, alcance a la izquierda y bono/castigo a la derecha */
function renderAdjustment(id, amount, rule, alcance, detail, emptyText) {
    const value = $(id);

    if (value) {
        value.textContent = formatSignedMoney(amount);
        value.classList.toggle("is-positive", amount > 0);
        value.classList.toggle("is-negative", amount < 0);
    }

    const effect = $(`${id}Effect`);

    if (!rule) {
        setText(`${id}Alcance`, "Alcance: —");
        setText(`${id}Note`, emptyText);

        if (effect) {
            effect.textContent = "";
            effect.className = "commission-effect";
        }

        return;
    }

    let text = "No se aplica descuento";
    let kind = "neutral";

    if (rule.pct > 0) {
        text = `Bono: +${rule.pct}%`;
        kind = "bonus";
    } else if (rule.pct < 0) {
        text = `Castigo: -${Math.abs(rule.pct)}%`;
        kind = "penalty";
    }

    setText(`${id}Alcance`, `Alcance: ${formatPercentage(alcance)}`);
    setText(`${id}Note`, detail);

    if (effect) {
        effect.textContent = text;
        effect.className = `commission-effect effect-${kind}`;
    }
}

function renderCommissionTable(currentRange, activeCount) {
    const tbody = $("commissionTableBody");

    if (!tbody) {
        return;
    }

    const nextRange = getNextCommissionRange(activeCount);

    tbody.innerHTML = COMMISSION_RANGES.map(range => {
        let rowClass = "";

        if (currentRange && range.range === currentRange.range) {
            rowClass = "current-range";
        } else if (nextRange && range.range === nextRange.range) {
            rowClass = "next-range";
        }

        return `
            <tr class="${rowClass}">
                <td>Rango ${range.range}</td>
                <td>${formatMoney(range.pos)}</td>
                <td>${formatMoney(range.alta)}</td>
                <td>${formatMoney(range.prepago)}</td>
                <td>${formatMoney(range.multi)}</td>
                <td>${range.from}</td>
                <td>${range.to === Infinity ? "∞" : range.to}</td>
            </tr>
        `;
    }).join("");
}


/* ============================================================
   DETALLE
============================================================ */

/* Acepta una venta o un índice dentro de filteredSales (compatibilidad) */
function mostrarDetalle(saleOrIndex) {
    const sale = typeof saleOrIndex === "number"
        ? filteredSales[saleOrIndex]
        : saleOrIndex;

    const body = $("detailBody");

    if (!sale || !body) {
        return;
    }

    const row = sale.row;

    setText("detailTitle", getValue(row, "nombre") || "Venta sin nombre");
    setText(
        "detailSubtitle",
        `Venta #${getValue(row, "n") || "—"} · DNI ${getValue(row, "dni") || "—"}`
    );

    body.innerHTML = DETAIL_SECTIONS.map(section => {
        const fieldsHTML = section.fields.map(([label, key, occurrence = 1]) => {
            let value = getValue(row, key, occurrence);

            if (DATE_KEYS.includes(key)) {
                value = formatDateValue(value);
            }

            return `
                <div class="detail-item">
                    <span class="detail-label">${escapeHTML(label)}</span>
                    <span class="detail-value">${value ? escapeHTML(value) : "—"}</span>
                </div>
            `;
        }).join("");

        return `
            <section class="detail-section">
                <h4 class="detail-section-title">${escapeHTML(section.title)}</h4>
                <div class="detail-grid">${fieldsHTML}</div>
            </section>
        `;
    }).join("");

    body.scrollTop = 0;

    abrirDetalle();
}

function abrirDetalle() {
    const overlay = $("detailOverlay");

    if (!overlay) {
        return;
    }

    lastFocusedElement = document.activeElement;

    overlay.hidden = false;
    document.body.style.overflow = "hidden";

    $("closeDetail")?.focus();
}

function cerrarDetalle() {
    const overlay = $("detailOverlay");

    if (!overlay || overlay.hidden) {
        return;
    }

    overlay.hidden = true;
    document.body.style.overflow = "";

    if (lastFocusedElement && document.contains(lastFocusedElement)) {
        lastFocusedElement.focus();
    }

    lastFocusedElement = null;
}

/* Mantiene el foco dentro del panel mientras está abierto */
function trapFocus(event) {
    const overlay = $("detailOverlay");

    if (event.key !== "Tab" || !overlay || overlay.hidden) {
        return;
    }

    const focusable = overlay.querySelectorAll(
        'button, [href], input, select, [tabindex]:not([tabindex="-1"])'
    );

    if (!focusable.length) {
        return;
    }

    const first = focusable[0];
    const last = focusable[focusable.length - 1];

    if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
    }
}


/* ============================================================
   TOAST
============================================================ */

function mostrarToast(message) {
    const toast = $("toast");

    if (!toast) {
        return;
    }

    toast.textContent = message;
    toast.hidden = false;

    clearTimeout(toastTimer);

    toastTimer = setTimeout(() => {
        toast.hidden = true;
    }, 4000);
}


/* ============================================================
   EVENTOS
============================================================ */

function configurarEventos() {
    const searchInput = $("searchInput");
    const advanced = $("advancedFilters");

    /* Espera a que el usuario termine de escribir */
    searchInput?.addEventListener("input", () => {
        clearTimeout(searchTimer);
        searchTimer = setTimeout(resetPageAndFilter, SEARCH_DEBOUNCE);
    });

    $("filtersToggle")?.addEventListener("click", () => {
        if ($("filtersPanel")?.hidden) {
            abrirFiltros();
        } else {
            cerrarFiltros();
        }
    });

    $("filtersApply")?.addEventListener("click", aplicarPanel);
    $("filtersCancel")?.addEventListener("click", cerrarFiltros);
    $("filtersReset")?.addEventListener("click", limpiarTodosPanel);

    advanced?.addEventListener("keydown", event => {
        if (event.key === "Enter" && event.target.matches("input")) {
            event.preventDefault();
            aplicarPanel();
        }
    });

    /* Clic fuera del panel = cancelar */
    document.addEventListener("click", event => {
        const menu = document.querySelector(".filters-menu");

        if (menu && !menu.contains(event.target)) {
            cerrarFiltros();
        }
    });

    $("clearFilters")?.addEventListener("click", () => {
        if (searchInput) searchInput.value = "";

        appliedValues = {};
        escribirControles({});

        cerrarFiltros();
        resetPageAndFilter();
    });

    $("refreshButton")?.addEventListener("click", () => cargarDatos(true));

    $("closeDetail")?.addEventListener("click", cerrarDetalle);

    $("detailOverlay")?.addEventListener("click", event => {
        if (event.target === event.currentTarget) {
            cerrarDetalle();
        }
    });

    document.addEventListener("keydown", event => {
        if (event.key === "Escape") {
            cerrarFiltros();
            cerrarDetalle();
        }

        trapFocus(event);
    });
}

function iniciarActualizacionAutomatica() {
    clearInterval(refreshTimer);

    refreshTimer = setInterval(() => {
        /* No consume datos si la pestaña está oculta */
        if (!document.hidden) {
            cargarDatos(false);
        }
    }, REFRESH_INTERVAL);

    /* Al volver a la pestaña, actualiza enseguida */
    document.addEventListener("visibilitychange", () => {
        if (!document.hidden) {
            cargarDatos(false);
        }
    });
}


/* ============================================================
   INICIO
============================================================ */

document.addEventListener("DOMContentLoaded", () => {
    construirFiltros();
    appliedValues = leerControles();
    configurarEventos();
    cargarDatos();
    iniciarActualizacionAutomatica();
});


/* ============================================================
   EXPORTACIONES
============================================================ */

window.mostrarDetalle = mostrarDetalle;
window.cerrarDetalle = cerrarDetalle;
window.cargarDatos = cargarDatos;
