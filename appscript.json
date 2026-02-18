/*******************************************************
 * SISTEMA COMPLETO:
 * - PEDIDOS_CONTROL (A:AF): ticket, fecha, timestamps, colores
 * - TICKET_PEDIDO: generar ticket manual (fila activa)
 * - REPORTE_DIARIO: reporte manual (hoy + pendientes acumulados)
 * - CATALOGO_PRODUCTOS (B:D): producto, stock inicial, activo
 * - ENTRADAS (A:F): registrar entradas manuales a KARDEX (sin duplicar)
 * - KARDEX: movimientos
 * - INVENTARIO_RESUMEN: existencias por producto/almacen
 * - TELEGRAM: enviar ticket a Telegram (fila activa)
 *
 * Reglas:
 * - Todo es ALMACEN_1
 * - Descuenta inventario cuando SALIO_A_REPARTO (AA) se marca TRUE
 * - Evita doble descuento por fila con PropertiesService
 *******************************************************/

// ===================== CONFIG PEDIDOS =====================
const CFG = {
  SHEET_CONTROL: "PEDIDOS_CONTROL",
  SHEET_TICKET: "TICKET_PEDIDO",
  SHEET_REPORTE: "REPORTE_DIARIO",
  TZ: "America/Mexico_City",

  // Columnas base
  COL_FECHA_PEDIDO: 1,   // A
  COL_TICKET: 2,         // B
  COL_CLIENTE: 3,        // C

  // Pares productos (PZ, PRODUCTO)
  PRODUCT_PAIRS: [
    { qty: 4,  prod: 5 },   // D-E
    { qty: 6,  prod: 7 },   // F-G
    { qty: 8,  prod: 9 },   // H-I
    { qty: 10, prod: 11 },  // J-K
    { qty: 12, prod: 13 },  // L-M
    { qty: 14, prod: 15 },  // N-O
    { qty: 16, prod: 17 },  // P-Q
    { qty: 18, prod: 19 },  // R-S
    { qty: 20, prod: 21 },  // T-U
    { qty: 22, prod: 23 }   // V-W
  ],

  COL_FOLIO: 24,        // X
  COL_FACTURADO: 25,    // Y (checkbox)
  COL_SURTIDO: 26,      // Z (checkbox)
  COL_SALIO: 27,        // AA (checkbox)

  COL_QUIEN: 28,        // AB
  COL_DOC: 29,          // AC
  COL_OBS: 30,          // AD
  COL_TS_FACT: 31,      // AE
  COL_TS_SALIO: 32      // AF
};

// ===================== CONFIG INVENTARIO =====================
const INV = {
  TZ: "America/Mexico_City",

  SHEET_PEDIDOS: "PEDIDOS_CONTROL",
  SHEET_CATALOGO: "CATALOGO_PRODUCTOS",
  SHEET_ENTRADAS: "ENTRADAS",
  SHEET_KARDEX: "KARDEX",
  SHEET_RESUMEN: "INVENTARIO_RESUMEN",

  // CATALOGO_PRODUCTOS (B:D)
  CAT_COL_PRODUCTO: 2,  // B
  CAT_COL_STOCK: 3,     // C
  CAT_COL_ACTIVO: 4,    // D (SI/NO)

  // KARDEX headers
  KARDEX_HEADERS: [
    "FECHA", "TIPO", "ALMACEN",
    "TICKET", "FOLIO", "CLIENTE",
    "PRODUCTO", "CANTIDAD",
    "QUIEN_LO_LLEVO", "DOCUMENTO_RECIBIDO",
    "OBSERVACION", "ID_MOV"
  ],

  ALMACEN_DEFAULT: "ALMACEN_1"
};

// ===================== CONFIG TELEGRAM =====================
const TG = {
  CHAT_ID: 7772460924 // tu chat id
};

// ===================== MENÚ ÚNICO =====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Pedidos")
    .addItem("🧾 Generar ticket (fila activa)", "generarTicketFilaActiva")
    .addItem("📨 Enviar ticket a Telegram (fila activa)", "enviarTicketTelegramFilaActiva")
    .addSeparator()
    .addItem("📊 Generar reporte de HOY + Pendientes", "generarReporteHoyConPendientes")
    .addSeparator()
    .addItem("🎨 Repintar fila activa", "repintarFilaActiva")
    .addToUi();

  ui.createMenu("Inventario")
    .addItem("🧾 Cargar stock inicial (CATALOGO → KARDEX)", "cargarStockInicial")
    .addItem("🔄 Recalcular INVENTARIO_RESUMEN", "recalcularInventarioResumen")
    .addToUi();

  ui.createMenu("Entradas")
    .addItem("📥 Registrar entrada (fila activa)", "registrarEntradaFilaActiva")
    .addItem("📥 Registrar entradas (selección)", "registrarEntradasSeleccion")
    .addToUi();

  ui.createMenu("Telegram")
    .addItem("🔐 Guardar/Actualizar token del bot", "setTelegramToken_")
    .addItem("🧪 Probar conexión (getMe)", "telegramTestGetMe_")
    .addToUi();
}

/**
 * IMPORTANTE:
 * Usa trigger instalable si quieres 100% confiable.
 * Pero dejamos también onEdit simple para que funcione sin trigger.
 */
function onEdit(e) {
  try { onEditHandler(e); } catch (err) {}
}

function onEditHandler(e) {
  if (!e) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEET_CONTROL) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  // 1) Asegurar fecha/ticket al capturar cliente o productos
  const isCliente = (col === CFG.COL_CLIENTE);
  const isAnyProductCol = CFG.PRODUCT_PAIRS.some(p => col === p.qty || col === p.prod);

  if (isCliente || isAnyProductCol) {
    asegurarFechaPedido_(sh, row);
    asegurarTicket_(sh, row);
  }

  // 2) Timestamp FACTURADO (Y->AE)
  if (col === CFG.COL_FACTURADO) {
    asegurarFechaPedido_(sh, row);
    asegurarTicket_(sh, row);

    const val = sh.getRange(row, CFG.COL_FACTURADO).getValue() === true;
    const tsCell = sh.getRange(row, CFG.COL_TS_FACT);
    if (val) {
      if (!tsCell.getValue()) tsCell.setValue(new Date());
    } else {
      tsCell.clearContent();
    }
  }

  // 3) Timestamp SALIO (AA->AF) + DESCUENTO INVENTARIO
  if (col === CFG.COL_SALIO) {
    asegurarFechaPedido_(sh, row);
    asegurarTicket_(sh, row);

    const val = sh.getRange(row, CFG.COL_SALIO).getValue() === true;

    const tsCell = sh.getRange(row, CFG.COL_TS_SALIO);
    if (val) {
      if (!tsCell.getValue()) tsCell.setValue(new Date());

      // descontar inventario + kardex
      procesarSalidaPedido_(row);

      // resumen
      recalcularInventarioResumen();
    }
  }

  // 4) Repintar fila por estado
  const relevantCols = [
    CFG.COL_CLIENTE,
    ...CFG.PRODUCT_PAIRS.flatMap(p => [p.qty, p.prod]),
    CFG.COL_FOLIO, CFG.COL_FACTURADO, CFG.COL_SURTIDO, CFG.COL_SALIO,
    CFG.COL_QUIEN, CFG.COL_DOC
  ];

  if (relevantCols.includes(col)) {
    pintarFilaSegunEstado_(sh, row);
    SpreadsheetApp.flush();
  }
}

// ===================== PEDIDOS HELPERS =====================
function asegurarFechaPedido_(sh, row) {
  const cell = sh.getRange(row, CFG.COL_FECHA_PEDIDO);
  if (!cell.getValue()) cell.setValue(new Date()).setNumberFormat("dd/MM/yyyy");
}

function asegurarTicket_(sh, row) {
  const cell = sh.getRange(row, CFG.COL_TICKET);
  if (cell.getValue()) return;

  const props = PropertiesService.getDocumentProperties();
  const last = Number(props.getProperty("LAST_TICKET") || "0");
  const next = last + 1;

  cell.setValue(next);
  props.setProperty("LAST_TICKET", String(next));
}

function tieneProductosEnFila_(sh, row) {
  for (const p of CFG.PRODUCT_PAIRS) {
    const qty = sh.getRange(row, p.qty).getValue();
    const prod = sh.getRange(row, p.prod).getValue();
    if ((qty && Number(qty) !== 0) || (prod && String(prod).trim() !== "")) return true;
  }
  return false;
}

function pintarFilaSegunEstado_(sh, row) {
  const fact = sh.getRange(row, CFG.COL_FACTURADO).getValue() === true;
  const surt = sh.getRange(row, CFG.COL_SURTIDO).getValue() === true;
  const salio = sh.getRange(row, CFG.COL_SALIO).getValue() === true;
  const doc = sh.getRange(row, CFG.COL_DOC).getValue();

  let color;
  if (doc && String(doc).trim() !== "") color = "#6aa84f";       // verde cerrado
  else if (salio) color = "#8e7cc3";                            // morado reparto
  else if (surt) color = "#6fa8dc";                             // azul surtido
  else if (fact) color = "#f6b26b";                             // naranja facturado
  else {
    const cliente = sh.getRange(row, CFG.COL_CLIENTE).getValue();
    const tieneAlgo = Boolean(cliente) || tieneProductosEnFila_(sh, row);
    color = tieneAlgo ? "#ffe599" : "#ffffff";                  // amarillo capturado
  }

  sh.getRange(row, 1, 1, 32).setBackground(color); // A:AF
}

function repintarFilaActiva() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_CONTROL);
  const r = ss.getActiveRange();
  if (!r || r.getSheet().getName() !== CFG.SHEET_CONTROL) {
    SpreadsheetApp.getUi().alert("Ponte en PEDIDOS_CONTROL y selecciona una celda de la fila.");
    return;
  }
  const row = r.getRow();
  if (row < 2) return;
  pintarFilaSegunEstado_(sh, row);
  SpreadsheetApp.flush();
}

// ===================== TICKET (MANUAL) =====================
function generarTicketFilaActiva() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_CONTROL);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_CONTROL);

  const active = ss.getActiveRange();
  if (!active || active.getSheet().getName() !== CFG.SHEET_CONTROL) {
    SpreadsheetApp.getUi().alert("Ponte en una fila de PEDIDOS_CONTROL y vuelve a intentar.");
    return;
  }

  const row = active.getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una fila de pedido (desde la fila 2).");
    return;
  }

  asegurarFechaPedido_(sh, row);
  asegurarTicket_(sh, row);

  const data = sh.getRange(row, 1, 1, 32).getValues()[0]; // A:AF
  const fecha = data[CFG.COL_FECHA_PEDIDO - 1];
  const ticket = data[CFG.COL_TICKET - 1];
  const cliente = data[CFG.COL_CLIENTE - 1];
  const folio = data[CFG.COL_FOLIO - 1] || "";
  const quien = data[CFG.COL_QUIEN - 1] || "";
  const obs = data[CFG.COL_OBS - 1] || "";

  const items = [];
  CFG.PRODUCT_PAIRS.forEach((p, idx) => {
    const qty = data[p.qty - 1];
    const prod = data[p.prod - 1];
    if ((qty && Number(qty) !== 0) || (prod && String(prod).trim() !== "")) {
      items.push([idx + 1, qty || "", prod || ""]);
    }
  });

  const t = ss.getSheetByName(CFG.SHEET_TICKET) || ss.insertSheet(CFG.SHEET_TICKET);
  t.clear();

  t.getRange("A1").setValue("TICKET DE PEDIDO").setFontSize(16).setFontWeight("bold");
  t.getRange("A3").setValue("TICKET:"); t.getRange("B3").setValue(ticket).setFontWeight("bold");
  t.getRange("A4").setValue("FECHA:");  t.getRange("B4").setValue(fecha).setNumberFormat("dd/MM/yyyy");
  t.getRange("A5").setValue("CLIENTE:");t.getRange("B5").setValue(cliente).setFontWeight("bold");
  t.getRange("A6").setValue("FOLIO FACTURA:"); t.getRange("B6").setValue(folio);
  t.getRange("A7").setValue("QUIÉN LO LLEVA:"); t.getRange("B7").setValue(quien);

  t.getRange("A9").setValue("PRODUCTOS").setFontWeight("bold");
  t.getRange("A10:C10").setValues([["#", "PZ", "PRODUCTO"]]).setFontWeight("bold");

  if (items.length) t.getRange(11, 1, items.length, 3).setValues(items);
  else t.getRange("A11").setValue("— Sin productos capturados —");

  const bottom = 13 + Math.max(items.length, 1);
  t.getRange(bottom, 1).setValue("OBSERVACIONES:").setFontWeight("bold");
  t.getRange(bottom + 1, 1).setValue(obs).setWrap(true);
  t.setRowHeights(bottom + 1, 1, 80);

  t.getRange(bottom + 3, 1).setValue("FIRMAS").setFontWeight("bold");
  t.getRange(bottom + 4, 1).setValue("BODEGA: ___________________________");
  t.getRange(bottom + 5, 1).setValue("REPARTO: ___________________________");
  t.getRange(bottom + 6, 1).setValue("CLIENTE: ___________________________");

  t.setColumnWidths(1, 1, 160);
  t.setColumnWidths(2, 1, 120);
  t.setColumnWidths(3, 1, 420);

  ss.setActiveSheet(t);
  SpreadsheetApp.getUi().alert("Ticket generado en 'TICKET_PEDIDO'. Imprime esa hoja.");
}

// ===================== REPORTE (HOY + PENDIENTES) =====================
function generarReporteHoyConPendientes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_CONTROL);
  if (!sh) throw new Error("No existe la hoja: " + CFG.SHEET_CONTROL);

  const rep = ss.getSheetByName(CFG.SHEET_REPORTE) || ss.insertSheet(CFG.SHEET_REPORTE);
  rep.clear();

  const today = new Date();
  const ymd = Utilities.formatDate(today, CFG.TZ, "yyyy-MM-dd");
  const fechaBonita = Utilities.formatDate(today, CFG.TZ, "dd/MM/yyyy");

  rep.getRange(1, 1).setValue("REPORTE - " + fechaBonita).setFontWeight("bold").setFontSize(14);
  rep.setFrozenRows(1);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rows = sh.getRange(2, 1, lastRow - 1, 32).getValues();

  const ssId = ss.getId();
  const gid = sh.getSheetId();

  const facturadosHoy = [];
  const salieronHoy = [];

  const pendientesFacturar = [];
  const pendientesSurtir = [];
  const pendientesReparto = [];

  rows.forEach((r, idx) => {
    const rowNum = idx + 2;
    const link = `https://docs.google.com/spreadsheets/d/${ssId}/edit#gid=${gid}&range=A${rowNum}`;

    const ticket = r[CFG.COL_TICKET - 1];
    const cliente = r[CFG.COL_CLIENTE - 1];
    const folio = r[CFG.COL_FOLIO - 1] || "";
    const quien = r[CFG.COL_QUIEN - 1] || "";
    const doc = r[CFG.COL_DOC - 1] || "";

    const fact = r[CFG.COL_FACTURADO - 1] === true;
    const surt = r[CFG.COL_SURTIDO - 1] === true;
    const salio = r[CFG.COL_SALIO - 1] === true;

    const tsFact = r[CFG.COL_TS_FACT - 1];
    const tsSal = r[CFG.COL_TS_SALIO - 1];

    const resumen = resumenProductos_(r);

    // HOY
    if (tsFact instanceof Date) {
      const ymdFact = Utilities.formatDate(tsFact, CFG.TZ, "yyyy-MM-dd");
      if (ymdFact === ymd) {
        facturadosHoy.push([ticket, folio, cliente, resumen, formatoHora_(tsFact), doc, `=HYPERLINK("${link}","Ver")`]);
      }
    }
    if (tsSal instanceof Date) {
      const ymdSal = Utilities.formatDate(tsSal, CFG.TZ, "yyyy-MM-dd");
      if (ymdSal === ymd) {
        salieronHoy.push([ticket, folio, cliente, quien, resumen, formatoHora_(tsSal), doc, `=HYPERLINK("${link}","Ver")`]);
      }
    }

    // PENDIENTES (acumulado)
    const tieneAlgo = Boolean(cliente) || resumen.length > 0;

    if (tieneAlgo) {
      if (!fact) pendientesFacturar.push([ticket, cliente, resumen, `=HYPERLINK("${link}","Ver")`]);
      if (fact && !surt) pendientesSurtir.push([ticket, folio, cliente, resumen, `=HYPERLINK("${link}","Ver")`]);
      if (surt && !salio) pendientesReparto.push([ticket, folio, cliente, resumen, `=HYPERLINK("${link}","Ver")`]);
    }
  });

  let r0 = 3;

  // FACTURADOS HOY
  rep.getRange(r0, 1).setValue("FACTURADOS HOY").setFontWeight("bold");
  rep.getRange(r0 + 1, 1, 1, 7).setValues([["TICKET","FACTURA","CLIENTE","PRODUCTOS","HORA","DOC_RECIBIDO","LINK"]]).setFontWeight("bold");
  if (facturadosHoy.length) rep.getRange(r0 + 2, 1, facturadosHoy.length, 7).setValues(facturadosHoy);
  else rep.getRange(r0 + 2, 1).setValue("— Sin facturados hoy —");

  r0 = r0 + 4 + Math.max(facturadosHoy.length, 1);

  // SALIERON HOY
  rep.getRange(r0, 1).setValue("SALIERON A REPARTO HOY").setFontWeight("bold");
  rep.getRange(r0 + 1, 1, 1, 8).setValues([["TICKET","FACTURA","CLIENTE","QUIÉN","PRODUCTOS","HORA","DOC_RECIBIDO","LINK"]]).setFontWeight("bold");
  if (salieronHoy.length) rep.getRange(r0 + 2, 1, salieronHoy.length, 8).setValues(salieronHoy);
  else rep.getRange(r0 + 2, 1).setValue("— Sin salidas hoy —");

  r0 = r0 + 4 + Math.max(salieronHoy.length, 1);

  // PENDIENTES
  rep.getRange(r0, 1).setValue("PENDIENTES ACUMULADOS").setFontWeight("bold");
  r0++;

  rep.getRange(r0, 1).setValue("Pendientes por FACTURAR").setFontWeight("bold");
  rep.getRange(r0 + 1, 1, 1, 4).setValues([["TICKET","CLIENTE","PRODUCTOS","LINK"]]).setFontWeight("bold");
  if (pendientesFacturar.length) rep.getRange(r0 + 2, 1, pendientesFacturar.length, 4).setValues(pendientesFacturar);
  else rep.getRange(r0 + 2, 1).setValue("— Ninguno —");

  r0 = r0 + 4 + Math.max(pendientesFacturar.length, 1);

  rep.getRange(r0, 1).setValue("Pendientes por SURTIR").setFontWeight("bold");
  rep.getRange(r0 + 1, 1, 1, 5).setValues([["TICKET","FACTURA","CLIENTE","PRODUCTOS","LINK"]]).setFontWeight("bold");
  if (pendientesSurtir.length) rep.getRange(r0 + 2, 1, pendientesSurtir.length, 5).setValues(pendientesSurtir);
  else rep.getRange(r0 + 2, 1).setValue("— Ninguno —");

  r0 = r0 + 4 + Math.max(pendientesSurtir.length, 1);

  rep.getRange(r0, 1).setValue("Pendientes por SALIR A REPARTO").setFontWeight("bold");
  rep.getRange(r0 + 1, 1, 1, 5).setValues([["TICKET","FACTURA","CLIENTE","PRODUCTOS","LINK"]]).setFontWeight("bold");
  if (pendientesReparto.length) rep.getRange(r0 + 2, 1, pendientesReparto.length, 5).setValues(pendientesReparto);
  else rep.getRange(r0 + 2, 1).setValue("— Ninguno —");

  rep.autoResizeColumns(1, 8);
}

function resumenProductos_(rowArray) {
  const parts = [];
  CFG.PRODUCT_PAIRS.forEach(p => {
    const qty = rowArray[p.qty - 1];
    const prod = rowArray[p.prod - 1];
    if ((qty && Number(qty) !== 0) || (prod && String(prod).trim() !== "")) {
      const q = qty ? qty : "";
      const pr = prod ? String(prod).trim() : "";
      parts.push(`${q} ${pr}`.trim());
    }
  });
  return parts.join(" | ");
}
function formatoHora_(d) {
  return Utilities.formatDate(d, CFG.TZ, "HH:mm");
}

// ===================== INVENTARIO: CARGA INICIAL =====================
function cargarStockInicial() {
  const ss = SpreadsheetApp.getActive();
  const cat = ss.getSheetByName(INV.SHEET_CATALOGO);
  if (!cat) throw new Error("No existe la hoja: " + INV.SHEET_CATALOGO);

  const kardex = obtenerOCrearKardex_();
  asegurarHeadersKardex_(kardex);

  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty("STOCK_INICIAL_CARGADO") === "1") {
    SpreadsheetApp.getUi().alert("⚠️ Ya cargaste el stock inicial antes. No se volverá a duplicar.");
    return;
  }

  const lastRow = cat.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("CATALOGO_PRODUCTOS está vacío.");
    return;
  }

  const data = cat.getRange(2, 2, lastRow - 1, 3).getValues(); // B:D

  const now = new Date();
  const movs = [];

  data.forEach(r => {
    const producto = (r[0] || "").toString().trim();
    const stock = Number(r[1] || 0);
    const activo = (r[2] || "").toString().trim().toUpperCase();

    if (!producto) return;
    if (activo !== "SI") return;
    if (!stock || stock === 0) return;

    movs.push([
      now, "ENTRADA_INICIAL", INV.ALMACEN_DEFAULT,
      "", "", "",
      producto, stock,
      "", "",
      "Carga inicial desde catálogo",
      "INIT-" + producto
    ]);
  });

  if (!movs.length) {
    SpreadsheetApp.getUi().alert("No encontré productos activos con stock > 0 para cargar.");
    return;
  }

  kardex.getRange(kardex.getLastRow() + 1, 1, movs.length, INV.KARDEX_HEADERS.length).setValues(movs);

  props.setProperty("STOCK_INICIAL_CARGADO", "1");

  recalcularInventarioResumen();
  SpreadsheetApp.getUi().alert("✅ Stock inicial cargado y resumen actualizado.");
}

// ===================== INVENTARIO: SALIDA POR PEDIDO =====================
function procesarSalidaPedido_(row) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(INV.SHEET_PEDIDOS);
  const kardex = obtenerOCrearKardex_();
  asegurarHeadersKardex_(kardex);

  const idMovKey = "PED_SALIO_ROW_" + row;
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(idMovKey) === "1") return;

  const r = sh.getRange(row, 1, 1, 32).getValues()[0]; // A:AF

  const ticket  = r[CFG.COL_TICKET - 1] || "";
  const cliente = r[CFG.COL_CLIENTE - 1] || "";
  const folio   = r[CFG.COL_FOLIO - 1] || "";
  const quien   = r[CFG.COL_QUIEN - 1] || "";
  const doc     = r[CFG.COL_DOC - 1] || "";
  const obs     = r[CFG.COL_OBS - 1] || "";

  const items = [];
  CFG.PRODUCT_PAIRS.forEach(p => {
    const qty = Number(r[p.qty - 1] || 0);
    const prod = (r[p.prod - 1] || "").toString().trim();
    if (prod && qty && qty !== 0) items.push({ prod, qty });
  });

  if (!items.length) {
    props.setProperty(idMovKey, "1");
    return;
  }

  const catalogSet = obtenerSetProductosCatalogo_();
  const now = new Date();
  const movs = [];

  items.forEach(it => {
    const producto = it.prod;
    const cantidad = it.qty;
    const notaExtra = catalogSet.has(producto) ? "" : " (⚠️ No está en CATALOGO_PRODUCTOS)";

    movs.push([
      now, "SALIDA_PEDIDO", INV.ALMACEN_DEFAULT,
      ticket, folio, cliente,
      producto, -Math.abs(cantidad),
      quien, doc,
      (obs ? obs : "Salida por pedido") + notaExtra,
      `PED-${row}-${producto}-${now.getTime()}`
    ]);
  });

  kardex.getRange(kardex.getLastRow() + 1, 1, movs.length, INV.KARDEX_HEADERS.length).setValues(movs);
  props.setProperty(idMovKey, "1");
}

// ===================== INVENTARIO: RESUMEN =====================
function recalcularInventarioResumen() {
  const ss = SpreadsheetApp.getActive();
  const kardex = ss.getSheetByName(INV.SHEET_KARDEX);
  const resumen = ss.getSheetByName(INV.SHEET_RESUMEN) || ss.insertSheet(INV.SHEET_RESUMEN);

  if (!kardex) {
    resumen.clear();
    resumen.getRange(1,1).setValue("No existe KARDEX");
    return;
  }

  asegurarHeadersKardex_(kardex);

  const lastRow = kardex.getLastRow();
  resumen.clear();
  resumen.getRange(1,1,1,5).setValues([["PRODUCTO","ALMACEN","EXISTENCIA","ENTRADAS","SALIDAS"]]).setFontWeight("bold");
  resumen.setFrozenRows(1);

  if (lastRow < 2) return;

  const data = kardex.getRange(2, 1, lastRow - 1, INV.KARDEX_HEADERS.length).getValues();
  const map = new Map();

  data.forEach(r => {
    const almacen = (r[2] || "").toString().trim() || INV.ALMACEN_DEFAULT;
    const producto = (r[6] || "").toString().trim();
    const qty = Number(r[7] || 0);
    if (!producto || !qty) return;

    const key = `${producto}||${almacen}`;
    if (!map.has(key)) map.set(key, { producto, almacen, existencia: 0, entradas: 0, salidas: 0 });

    const obj = map.get(key);
    obj.existencia += qty;
    if (qty > 0) obj.entradas += qty;
    if (qty < 0) obj.salidas += Math.abs(qty);
  });

  const rows = Array.from(map.values())
    .sort((a,b) => a.producto.localeCompare(b.producto))
    .map(x => [x.producto, x.almacen, x.existencia, x.entradas, x.salidas]);

  if (rows.length) {
    resumen.getRange(2, 1, rows.length, 5).setValues(rows);
    resumen.autoResizeColumns(1, 5);
  }
}

// ===================== ENTRADAS: REGISTRAR MANUAL (SIN DUPLICAR) =====================
// ENTRADAS: A FECHA | B ALMACEN | C PRODUCTO | D CANTIDAD | E PROVEEDOR | F OBS

function registrarEntradaFilaActiva() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  if (sh.getName() !== INV.SHEET_ENTRADAS) {
    SpreadsheetApp.getUi().alert("Ponte en la hoja ENTRADAS y selecciona una fila.");
    return;
  }

  const row = ss.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una fila válida (desde la fila 2).");
    return;
  }

  const res = registrarEntradaPorFila_(row);
  SpreadsheetApp.getUi().alert(res);
}

function registrarEntradasSeleccion() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  if (sh.getName() !== INV.SHEET_ENTRADAS) {
    SpreadsheetApp.getUi().alert("Ponte en la hoja ENTRADAS y selecciona filas.");
    return;
  }

  const r = ss.getActiveRange();
  const startRow = r.getRow();
  const numRows = r.getNumRows();

  if (startRow < 2) {
    SpreadsheetApp.getUi().alert("Selecciona filas desde la fila 2.");
    return;
  }

  let ok = 0, skip = 0, err = 0;

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const msg = registrarEntradaPorFila_(row, true);
    if (msg.startsWith("✅")) ok++;
    else if (msg.startsWith("⏭️")) skip++;
    else err++;
  }

  SpreadsheetApp.getUi().alert(`Listo.\n✅ Registradas: ${ok}\n⏭️ Omitidas: ${skip}\n❌ Error: ${err}`);
}

function registrarEntradaPorFila_(row) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(INV.SHEET_ENTRADAS);
  const kardex = obtenerOCrearKardex_();
  asegurarHeadersKardex_(kardex);

  const props = PropertiesService.getDocumentProperties();
  const key = `ENTRADA_ROW_${row}`;

  if (props.getProperty(key) === "1") {
    return "⏭️ Ya estaba registrada (fila " + row + ")";
  }

  const vals = sh.getRange(row, 1, 1, 6).getValues()[0];
  const fecha = vals[0];
  const almacenRaw = vals[1];
  const productoRaw = vals[2];
  const cantidadRaw = vals[3];
  const proveedor = vals[4];
  const obs = vals[5];

  const producto = (productoRaw || "").toString().trim();
  const cantidad = Number(cantidadRaw || 0);

  if (!producto || !cantidad || cantidad === 0) {
  return "⏭️ Fila " + row + " vacía o cantidad inválida";
}

  // Regla: todo ALMACEN_1 (aceptamos "1" o vacío)
  let almacen = (almacenRaw || "").toString().trim();
  if (!almacen || almacen === "1") almacen = INV.ALMACEN_DEFAULT;

  const f = (fecha instanceof Date) ? fecha : new Date();

  kardex.appendRow([
    f,
    "ENTRADA_COMPRA",
    almacen,
    "", "", "",
    producto,
    cantidad,
    "",
    proveedor || "",
    obs || "Entrada manual",
    `ENT-${row}-${new Date().getTime()}`
  ]);

  props.setProperty(key, "1");

  // Morado = entrada registrada
  sh.getRange(row, 1, 1, 6).setBackground("#8e7cc3");

  recalcularInventarioResumen();
  return "✅ Entrada registrada (fila " + row + ")";
}

// ===================== HELPERS KARDEX/CATALOGO =====================
function obtenerOCrearKardex_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(INV.SHEET_KARDEX) || ss.insertSheet(INV.SHEET_KARDEX);
}

function asegurarHeadersKardex_(kardex) {
  const firstRow = kardex.getRange(1,1,1,INV.KARDEX_HEADERS.length).getValues()[0];
  const ok = INV.KARDEX_HEADERS.every((h,i) => (firstRow[i] || "").toString().trim() === h);
  if (!ok) {
    kardex.getRange(1,1,1,INV.KARDEX_HEADERS.length).setValues([INV.KARDEX_HEADERS]).setFontWeight("bold");
    kardex.setFrozenRows(1);
  }
}

function obtenerSetProductosCatalogo_() {
  const ss = SpreadsheetApp.getActive();
  const cat = ss.getSheetByName(INV.SHEET_CATALOGO);
  if (!cat) return new Set();

  const last = cat.getLastRow();
  if (last < 2) return new Set();

  const data = cat.getRange(2, 2, last - 1, 3).getValues(); // B:D
  const set = new Set();

  data.forEach(r => {
    const prod = (r[0] || "").toString().trim();
    const activo = (r[2] || "").toString().trim().toUpperCase();
    if (prod && activo === "SI") set.add(prod);
  });

  return set;
}

// ===================== TELEGRAM: TOKEN + ENVIAR TICKET =====================

function setTelegramToken_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Guardar token de Telegram",
    "Pega aquí el token EXACTO de BotFather (sin espacios):",
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const token = String(resp.getResponseText() || "").trim();

  if (!token || token.indexOf(":") === -1) {
    ui.alert("❌ Token inválido. Debe verse como: 123456789:AA...");
    return;
  }

  PropertiesService.getDocumentProperties().setProperty("TELEGRAM_BOT_TOKEN", token);
  ui.alert("✅ Token guardado correctamente.");
}

function enviarTicketTelegramFilaActiva() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEET_CONTROL);
  const r = ss.getActiveRange();

  if (!r || r.getSheet().getName() !== CFG.SHEET_CONTROL) {
    SpreadsheetApp.getUi().alert("Ponte en PEDIDOS_CONTROL y selecciona una celda de la fila del pedido.");
    return;
  }

  const row = r.getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una fila válida (desde la fila 2).");
    return;
  }

  // Asegura fecha y ticket
  asegurarFechaPedido_(sh, row);
  asegurarTicket_(sh, row);

  const data = sh.getRange(row, 1, 1, 32).getValues()[0]; // A:AF

  const fecha   = data[CFG.COL_FECHA_PEDIDO - 1];
  const ticket  = data[CFG.COL_TICKET - 1];
  const cliente = data[CFG.COL_CLIENTE - 1] || "";
  const folio   = data[CFG.COL_FOLIO - 1] || "";
  const quien   = data[CFG.COL_QUIEN - 1] || "";
  const doc     = data[CFG.COL_DOC - 1] || "";
  const obs     = data[CFG.COL_OBS - 1] || "";

  const fact = data[CFG.COL_FACTURADO - 1] === true;
  const surt = data[CFG.COL_SURTIDO - 1] === true;
  const salio = data[CFG.COL_SALIO - 1] === true;

  const fechaTxt = (fecha instanceof Date)
    ? Utilities.formatDate(fecha, CFG.TZ, "dd/MM/yyyy")
    : (fecha ? String(fecha) : "");

  const items = [];
  CFG.PRODUCT_PAIRS.forEach(p => {
    const qty = Number(data[p.qty - 1] || 0);
    const prod = (data[p.prod - 1] || "").toString().trim();
    if (prod && qty && qty !== 0) items.push({ qty, prod });
  });

  let text = "";
  text += `🧾 *TICKET ${escapeMd_(ticket)}*\n`;
  if (fechaTxt) text += `📅 ${escapeMd_(fechaTxt)}\n`;
  text += `👤 *${escapeMd_(cliente)}*\n`;
  if (folio) text += `🧾 Factura: ${escapeMd_(folio)}\n`;

  text += `\n`;
  if (items.length) {
    items.forEach(it => text += `• ${it.qty}  ${escapeMd_(it.prod)}\n`);
  } else {
    text += "— Sin productos —\n";
  }

  text += `\n📌 Estado: ${fact ? "FACTURADO ✅" : "FACTURADO ❌"} | ${surt ? "SURTIDO ✅" : "SURTIDO ❌"} | ${salio ? "REPARTO ✅" : "REPARTO ❌"}\n`;

  if (quien) text += `🚚 Quién: ${escapeMd_(quien)}\n`;
  if (doc) text += `📄 Doc: ${escapeMd_(doc)}\n`;
  if (obs) text += `📝 Obs: ${escapeMd_(obs)}\n`;

  sendTelegramMessage_(TG.CHAT_ID, text);
  SpreadsheetApp.getUi().alert("✅ Ticket enviado a Telegram.");
}

function sendTelegramMessage_(chatId, text) {
  const token = PropertiesService.getDocumentProperties().getProperty("TELEGRAM_BOT_TOKEN");
  if (!token) {
    SpreadsheetApp.getUi().alert("Primero guarda tu token: Menú Telegram → Guardar/Actualizar token");
    throw new Error("No hay token guardado en PropertiesService.");
  }

  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: text,
    parse_mode: "Markdown"
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const out = JSON.parse(res.getContentText());
  if (!out.ok) throw new Error("Telegram error: " + res.getContentText());
}

function escapeMd_(s) {
  return String(s).replace(/([_*[\]()~`>#+\-=|{}.!\\])/g, "\\$1");
}

function telegramTestGetMe_() {
  const token = PropertiesService.getDocumentProperties().getProperty("TELEGRAM_BOT_TOKEN");
  if (!token) throw new Error("No hay token guardado. Usa Menú Telegram → Guardar/Actualizar token.");

  const url = "https://api.telegram.org/bot" + token + "/getMe";
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const txt = res.getContentText();

  Logger.log(txt);

  const json = JSON.parse(txt);
  if (!json.ok) throw new Error("Telegram error: " + txt);

  SpreadsheetApp.getUi().alert("🤖 Bot conectado: @" + json.result.username);
}

