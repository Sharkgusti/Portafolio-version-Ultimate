// =================================================================================
// ===   TITANIUM v2 — FISCAL.GS                                                ===
// ===   Módulo Bienes Personales (AFIP/ARCA)                                   ===
// ===   FIX: tenencia de USD/CASH lee de Monto_Neto_USD (col K), no Unidades   ===
// =================================================================================


const FISCAL_PARAM_SHEET  = 'PARAM_FISCAL';
const FISCAL_OUTPUT_SHEET = 'FISCAL_BS_PERSONALES';
const FISCAL_ARCA_SHEET   = 'ARCA_FUENTES';

const FISCAL_ANO_MIN = 2024;
const FISCAL_ANO_MAX = 2038;

const FISCAL_MANUAL_HEADER_ROW = 114;
const FISCAL_MANUAL_FIRST_ROW  = 115;
const FISCAL_MANUAL_LAST_ROW   = 130;

// =================================================================================
// HELPER: Sanitización numérica robusta para columna "Cotización" de ARCA_FUENTES
// =================================================================================
function sanitizarCotizacionARCA(celda) {
  if (typeof celda === 'number') return isFinite(celda) ? celda : null;
  if (celda === null || celda === undefined) return null;

  let s = String(celda).trim();
  if (!s) return null;

  const tienePunto = s.includes('.');
  const tieneComa  = s.includes(',');

  if (tieneComa && !tienePunto) {
    s = s.replace(',', '.');
  } else if (tienePunto && tieneComa) {
    s = s.replace(/\./g, '').replace(',', '.');
  }

  const num = parseFloat(s);
  return isNaN(num) ? null : num;
}

// =================================================================================
// HELPER: Índice de cotizaciones oficiales ARCA para el año fiscal
// =================================================================================
function construirIndiceArcaFuentes(ss, anoFiscal) {
  const indice = {};
  const arcaSheet = ss.getSheetByName(FISCAL_ARCA_SHEET);
  if (!arcaSheet) return indice;

  const lastRow = arcaSheet.getLastRow();
  if (lastRow < 2) return indice;

  const data = arcaSheet.getRange(2, 1, lastRow - 1, 13).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const anoFila = parseInt(row[0]);
    if (anoFila !== anoFiscal) continue;

    const ticker = String(row[3]).toUpperCase().trim();
    if (!ticker) continue;

    const cotizacionSan = sanitizarCotizacionARCA(row[12]);
    if (cotizacionSan === null || cotizacionSan <= 0) continue;

    indice[ticker] = { precioRefARS: cotizacionSan, filaOriginal: i + 2 };
  }

  return indice;
}

// =================================================================================
// HELPER: Lee el bloque "Otros Bienes — Carga Manual" (PARAM_FISCAL)
// =================================================================================
function leerBienesManualesParamFiscal(paramSheet) {
  const items = [];
  const nFilas = FISCAL_MANUAL_LAST_ROW - FISCAL_MANUAL_FIRST_ROW + 1;
  const data = paramSheet.getRange(FISCAL_MANUAL_FIRST_ROW, 1, nFilas, 4).getValues();

  for (const row of data) {
    const descripcion = String(row[0]).trim();
    const valorOriginal = cleanNum(row[1]);
    const moneda = String(row[2]).toUpperCase().trim();
    const condicion = String(row[3]).toUpperCase().trim();

    if (!descripcion) continue;
    if (valorOriginal <= 0) continue;
    if (!['USD', 'ARS'].includes(moneda)) continue;
    if (!['GRAVADO', 'EXENTO', 'NO COMPUTABLE'].includes(condicion)) continue;

    items.push({ descripcion, valorOriginal, moneda, condicion });
  }

  return items;
}

// =================================================================================
// FUNCIÓN PRINCIPAL — Genera el Informe de Bienes Personales
// =================================================================================
function GENERAR_INFORME_BIENES_PERSONALES() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── 1. PARÁMETROS ──────────────────────────────────────────────────
  const paramSheet = ss.getSheetByName(FISCAL_PARAM_SHEET);
  if (!paramSheet) {
    ui.alert('❌ Error: No existe la hoja PARAM_FISCAL.');
    return;
  }

  const anoFiscal = parseInt(paramSheet.getRange('B1').getValue());
  const tcBNA     = parseFloat(paramSheet.getRange('B2').getValue());

  if (!anoFiscal || anoFiscal < FISCAL_ANO_MIN || anoFiscal > FISCAL_ANO_MAX) {
    ui.alert('❌ Error: Seleccioná un año fiscal válido en PARAM_FISCAL!B1 (entre ' + FISCAL_ANO_MIN + ' y ' + FISCAL_ANO_MAX + ').');
    return;
  }
  if (!tcBNA || tcBNA < 100) {
    ui.alert('❌ Error: Ingresá el TC BNA Comprador al 31/12/' + anoFiscal + ' en PARAM_FISCAL!B2\n(valor actual: ' + tcBNA + ')');
    return;
  }

  const dictRaw = paramSheet.getRange('A6:B100').getValues();
  const diccionario = {};
  for (const row of dictRaw) {
    const tk  = String(row[0]).toUpperCase().trim();
    const cond = String(row[1]).toUpperCase().trim();
    if (tk && ['GRAVADO', 'EXENTO', 'NO COMPUTABLE'].includes(cond)) diccionario[tk] = cond;
  }

  const mniRaw = paramSheet.getRange('D6:F' + (5 + (FISCAL_ANO_MAX - FISCAL_ANO_MIN + 1))).getValues();
  let mniAno = 0, mniDefinido = false;
  for (const row of mniRaw) {
    if (parseInt(row[0]) === anoFiscal) {
      mniAno = cleanNum(row[1]);
      mniDefinido = (row[1] !== '' && row[1] !== null && !isNaN(parseFloat(row[1])));
      break;
    }
  }

  const escalaRaw = paramSheet.getRange('H6:L' + (5 + (FISCAL_ANO_MAX - FISCAL_ANO_MIN + 1) + 3)).getValues();
  const escalas = [];
  let hayEscalaParaElAno = false;
  let escalaPendienteCompletar = false;
  for (const row of escalaRaw) {
    if (parseInt(row[0]) === anoFiscal) {
      hayEscalaParaElAno = true;
      const alicuotaRaw = row[4];
      if (alicuotaRaw === '' || alicuotaRaw === null || isNaN(parseFloat(alicuotaRaw))) {
        escalaPendienteCompletar = true;
        continue;
      }
      escalas.push({
        desde: cleanNum(row[1]),
        hasta: (cleanNum(row[2]) === -1) ? Infinity : cleanNum(row[2]),
        fija: cleanNum(row[3]),
        alicuota: cleanNum(row[4]) / 100
      });
    }
  }

  // ── 2. TIMESTAMP DE CORTE ─────────────────────────────────────────
  const fechaCorte = new Date(anoFiscal, 11, 31, 23, 59, 59);

  // ── 3. ÍNDICE ARCA ────────────────────────────────────────────────
  const indiceArca = construirIndiceArcaFuentes(ss, anoFiscal);

  // ── 4. PROCESAR LOG — inventario al cierre ───────────────────────
  const logSheet = ss.getSheetByName('Log_Transacciones_TITANIUM');
  if (!logSheet) { ui.alert('❌ Error: No se encuentra Log_Transacciones_TITANIUM.'); return; }

  const logData = logSheet.getDataRange().getValues();
  const inventario = {}; // { ticker: { qty, tipo, moneda } }

  for (let i = 1; i < logData.length; i++) {
    const row  = logData[i];
    const fecha = new Date(row[1]);
    if (!(fecha instanceof Date) || isNaN(fecha) || fecha > fechaCorte) continue;

    const ticker  = String(row[3]).toUpperCase().trim();
    const tipoStr = String(row[4]);
    const mov     = String(row[5]).toLowerCase().trim();
    const cant    = cleanNum(row[6]);
    const montoUSD = Math.abs(cleanNum(row[10])); // Monto_Neto_USD, columna K (índice 10)

    const colH = String(row[7]).toUpperCase().trim();
    const colI = String(row[8]).toUpperCase().trim();
    const moneda = (colH === 'USD' || colI === 'USD') ? 'USD' : 'ARS';

    if (!ticker) continue;
    if (!inventario[ticker]) inventario[ticker] = { qty: 0, tipo: tipoStr, moneda };

    // =================================================================
    // FIX: para USD/CASH, la "cantidad" que importa es el monto en USD
    // (columna K), NO las Unidades (columna G) — que quedan vacías
    // cuando el aporte se cargó en pesos.
    // =================================================================
    if (ticker === 'USD' || ticker === 'CASH') {
      if (mov.includes('aporte') || mov.includes('compra') || mov.includes('suscripcion') || mov.includes('canje_entrada')) {
        inventario[ticker].qty += montoUSD;
      } else if (mov.includes('retiro') || mov.includes('venta') || mov.includes('rescate') || mov.includes('canje_salida')) {
        inventario[ticker].qty -= montoUSD;
      }
      continue;
    }

    // Resto de los activos: se mantiene la lógica original (Unidades)
    if (mov.includes('compra') || mov.includes('aporte') || mov.includes('suscripcion') || mov.includes('canje_entrada')) {
      inventario[ticker].qty += cant;
    } else if (mov.includes('venta') || mov.includes('rescate') || mov.includes('canje_salida')) {
      inventario[ticker].qty -= cant;
    } else if (mov.includes('split')) {
      const ratio = cleanNum(row[11]);
      if (ratio > 0) inventario[ticker].qty *= ratio;
    }
  }

  // ── 5. FALLBACK — Historico_Precios ───────────────────────────────
  const histSheet = ss.getSheetByName('Historico_Precios');
  const histMap = {};

  if (histSheet) {
    const histData = histSheet.getDataRange().getValues().slice(1);
    for (const row of histData) {
      if (!(row[0] instanceof Date) || isNaN(row[0])) continue;
      const tk     = String(row[1]).toUpperCase().trim();
      const precio = cleanNum(row[2]);
      if (!tk || precio <= 0) continue;
      if (!histMap[tk]) histMap[tk] = [];
      histMap[tk].push({ fecha: row[0], precioUSD: precio });
    }
  }

  function getPrecioUSDalCorte(ticker) {
    const puntos = histMap[ticker];
    if (!puntos || puntos.length === 0) return null;
    const candidatos = puntos.filter(p => p.fecha <= fechaCorte); // sin lookahead bias
    if (candidatos.length === 0) return null;
    candidatos.sort((a, b) => b.fecha - a.fecha);
    return candidatos[0].precioUSD;
  }

  // ── 6. VALUACIÓN — activos de bróker ──────────────────────────────
  const filasReporte = [];
  let hayPreciosFaltantes = [];

  for (const tk of Object.keys(inventario)) {
    const item = inventario[tk];
    if (item.qty < 0.001) continue;

    let valorARS = 0, precioRefARS = 0, notaPrecio = '', fuente = '';

    if (tk === 'USD' || tk === 'CASH' || tk === 'LIQUIDEZ_AUTO') {
      precioRefARS = tcBNA;
      valorARS = item.qty * tcBNA;
      fuente = 'TC Directo';
    } else if (indiceArca[tk]) {
      precioRefARS = indiceArca[tk].precioRefARS;
      valorARS = item.qty * precioRefARS;
      fuente = 'ARCA Oficial';
    } else {
      const precioUSD = getPrecioUSDalCorte(tk);
      if (precioUSD === null) {
        hayPreciosFaltantes.push(tk);
        precioRefARS = 0;
        valorARS = 0;
        notaPrecio = '⚠️ Sin precio histórico ni dato ARCA';
        fuente = 'Sin dato';
      } else {
        precioRefARS = precioUSD * tcBNA;
        valorARS = item.qty * precioUSD * tcBNA;
        fuente = 'Histórico + TC';
      }
    }

    const condicion  = diccionario[tk] || 'GRAVADO';
    const computable = (condicion === 'GRAVADO') ? valorARS : 0;
    const tipoNorm   = normalizar(item.tipo);

    filasReporte.push({
      ticker: tk, tipo: tipoNorm, qty: item.qty,
      precioARS: precioRefARS, valorARS: valorARS,
      condicion: condicion, computable: computable,
      fuente: fuente, nota: notaPrecio
    });
  }

  // ── 7. BIENES FUERA DE BRÓKER ──────────────────────────────────────
  const bienesManuales = leerBienesManualesParamFiscal(paramSheet);
  for (const item of bienesManuales) {
    const valorARS = (item.moneda === 'USD') ? item.valorOriginal * tcBNA : item.valorOriginal;
    const computable = (item.condicion === 'GRAVADO') ? valorARS : 0;

    filasReporte.push({
      ticker: item.descripcion, tipo: 'Manual', qty: item.valorOriginal,
      precioARS: (item.moneda === 'USD') ? tcBNA : 1,
      valorARS: valorARS, condicion: item.condicion, computable: computable,
      fuente: (item.moneda === 'USD') ? 'Manual (USD→ARS por TC)' : 'Manual (ARS)',
      nota: ''
    });
  }

  filasReporte.sort((a, b) => b.valorARS - a.valorARS);

  // ── 8. CÁLCULO FISCAL ──────────────────────────────────────────────
  const totalGravado  = filasReporte.reduce((s, r) => s + r.computable, 0);
  const baseImponible = Math.max(0, totalGravado - mniAno);

  let impuesto = 0, tramoDesc = '', alicuotaEfectiva = 0;
  let avisoEscalaPendiente = false, avisoMniPendiente = false;

  if (!mniDefinido) {
    avisoMniPendiente = true;
    tramoDesc = 'MNI no definido para ' + anoFiscal;
  } else if (baseImponible <= 0) {
    tramoDesc = 'No alcanza el MNI';
  } else if (escalas.length === 0) {
    if (escalaPendienteCompletar || !hayEscalaParaElAno) {
      avisoEscalaPendiente = true;
      tramoDesc = 'Escala de alícuotas no definida para ' + anoFiscal;
    } else {
      tramoDesc = 'Sin escala para el año';
    }
  } else {
    for (let i = escalas.length - 1; i >= 0; i--) {
      const tramo = escalas[i];
      if (baseImponible > tramo.desde) {
        impuesto = tramo.fija + (baseImponible - tramo.desde) * tramo.alicuota;
        tramoDesc = (tramo.alicuota * 100).toFixed(2) + '%';
        alicuotaEfectiva = impuesto / baseImponible;
        break;
      }
    }
    if (tramoDesc === '') tramoDesc = 'Sin tramo aplicable';
  }

  // ── 9. ESCRIBIR REPORTE ────────────────────────────────────────────
  let out = ss.getSheetByName(FISCAL_OUTPUT_SHEET);
  if (!out) out = ss.insertSheet(FISCAL_OUTPUT_SHEET);
  out.clear();
  out.clearFormats();

  const AZUL_OSCURO = '#0a192f', DORADO = '#FFC300', AZUL_MEDIO = '#1a3a5c', BLANCO = '#FFFFFF';
  const N_COLS = 9;
  let f = 1;

  out.getRange(f, 1, 1, N_COLS).merge()
    .setValue('INFORME DE AUDITORÍA FISCAL: BIENES PERSONALES — PERÍODO FISCAL ' + anoFiscal)
    .setFontWeight('bold').setFontSize(14).setBackground(AZUL_OSCURO).setFontColor(DORADO)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  out.setRowHeight(f, 36);
  f++;

  out.getRange(f, 1, 1, N_COLS).merge()
    .setValue('Foto patrimonial al 31/12/' + anoFiscal + '  |  TC BNA Comprador (fallback): $' + tcBNA.toLocaleString('es-AR', { minimumFractionDigits: 2 }) + '  |  Generado: ' + new Date().toLocaleDateString('es-AR'))
    .setFontStyle('italic').setFontSize(10).setBackground('#e8eaf6').setHorizontalAlignment('center');
  f++; f++;

  const headers = ['Ticker / Bien', 'Tipo de Activo', 'Cant. Nominal / Monto Orig.', 'Precio Ref. ARS', 'Valor Total ARS', 'Condición Impositiva', 'Computable ARS', 'Fuente del Precio', 'Nota'];
  out.getRange(f, 1, 1, N_COLS).setValues([headers])
    .setFontWeight('bold').setBackground(AZUL_MEDIO).setFontColor(BLANCO).setHorizontalAlignment('center');
  f++;

  for (let i = 0; i < filasReporte.length; i++) {
    const r = filasReporte[i];
    out.getRange(f, 1, 1, N_COLS).setValues([[r.ticker, r.tipo, r.qty, r.precioARS, r.valorARS, r.condicion, r.computable, r.fuente, r.nota]]);

    out.getRange(f, 3).setNumberFormat('#,##0.0000');
    out.getRange(f, 4).setNumberFormat('$#,##0.00');
    out.getRange(f, 5).setNumberFormat('$#,##0.00');
    out.getRange(f, 7).setNumberFormat('$#,##0.00');

    const bg = (i % 2 === 0) ? '#f5f7fa' : BLANCO;
    out.getRange(f, 1, 1, N_COLS).setBackground(bg);

    const celdaCond = out.getRange(f, 6);
    if (r.condicion === 'EXENTO') celdaCond.setFontColor('#2e7d32').setFontWeight('bold');
    else if (r.condicion === 'NO COMPUTABLE') celdaCond.setFontColor('#1565c0').setFontWeight('bold');
    else celdaCond.setFontColor('#c62828').setFontWeight('bold');

    const celdaFuente = out.getRange(f, 8);
    if (r.fuente === 'ARCA Oficial') celdaFuente.setFontColor('#2e7d32').setFontWeight('bold');
    else if (r.fuente.indexOf('Manual') === 0) celdaFuente.setFontColor('#6a1b9a');
    else if (r.fuente === 'Sin dato') celdaFuente.setFontColor('#c62828').setFontWeight('bold');
    else celdaFuente.setFontColor('#1565c0');

    f++;
  }

  f++;

  out.getRange(f, 1, 1, N_COLS).merge()
    .setValue('CUADRO DE LIQUIDACIÓN IMPOSITIVA')
    .setFontWeight('bold').setBackground(AZUL_MEDIO).setFontColor(DORADO).setHorizontalAlignment('center');
  f++;

  const resumenItems = [
    { label: '(+)   Total de Bienes Gravados Computables:', valor: totalGravado, fmt: '$#,##0.00', bold: false },
    { label: '(-)   Mínimo No Imponible Aplicado (ARCA):', valor: mniDefinido ? mniAno : 'No definido', fmt: mniDefinido ? '$#,##0.00' : null, bold: false },
    { label: '(=)   Base Imponible Sujeta a Impuesto:', valor: baseImponible, fmt: '$#,##0.00', bold: true },
    { label: '(*)   Tramo / Alícuota Marginal Aplicada:', valor: tramoDesc, fmt: null, bold: false },
    { label: '(*)   Alícuota Efectiva Total:', valor: alicuotaEfectiva * 100, fmt: '0.00"%"', bold: false },
    { label: '(=)   IMPUESTO ANUAL ESTIMADO A PAGAR:', valor: impuesto, fmt: '$#,##0.00', bold: true }
  ];

  for (const item of resumenItems) {
    const isFinalRow = (item.label.includes('IMPUESTO ANUAL'));
    const celdaLabel = out.getRange(f, 3);
    const celdaValor = out.getRange(f, 4);

    celdaLabel.setValue(item.label).setFontWeight(item.bold ? 'bold' : 'normal');
    celdaValor.setValue(item.valor);
    if (item.fmt) celdaValor.setNumberFormat(item.fmt);

    if (isFinalRow) {
      out.getRange(f, 3, 1, 2).setFontSize(12).setFontWeight('bold')
        .setBackground(impuesto > 0 ? '#ffebee' : '#e8f5e9');
      celdaValor.setFontColor(impuesto > 0 ? '#c62828' : '#2e7d32');
    }
    f++;
  }

  f++;

  if (avisoMniPendiente || avisoEscalaPendiente) {
    let msgPend = '⚠️ PARAMETRIZACIÓN PENDIENTE PARA ' + anoFiscal + ': ';
    const partes = [];
    if (avisoMniPendiente) partes.push('falta el Mínimo No Imponible (PARAM_FISCAL, Tabla 2)');
    if (avisoEscalaPendiente) partes.push('falta la Escala de Alícuotas (PARAM_FISCAL, Tabla 3)');
    msgPend += partes.join(' y ') + '. El impuesto NO pudo calcularse hasta completar estos datos.';

    out.getRange(f, 1, 1, N_COLS).merge().setValue(msgPend)
      .setFontWeight('bold').setFontStyle('italic').setFontSize(9)
      .setBackground('#fff3e0').setFontColor('#e65100').setWrap(true);
    out.setRowHeight(f, 40);
    f++;
  }

  const nota1 = '⚠️  AVISO LEGAL: Este cálculo es una ESTIMACIÓN educativa. No reemplaza el asesoramiento de un Contador Público (CPCE). ' +
                'Las cotizaciones marcadas como "ARCA Oficial" provienen de ARCA_FUENTES y se toman como valor en pesos sin conversión adicional.';
  out.getRange(f, 1, 1, N_COLS).merge().setValue(nota1)
    .setFontStyle('italic').setFontSize(9).setBackground('#fff8e1').setFontColor('#5d4037').setWrap(true);
  out.setRowHeight(f, 50);
  f++;

  if (hayPreciosFaltantes.length > 0) {
    out.getRange(f, 1, 1, N_COLS).merge()
      .setValue('⚠️ TICKERS SIN DATOS (ni ARCA ni histórico — valor forzado a $0, revisar manualmente): ' + hayPreciosFaltantes.join(', '))
      .setFontStyle('italic').setFontSize(9).setBackground('#fce4ec').setFontColor('#c62828').setWrap(true);
    f++;
  }

  out.autoResizeColumns(1, N_COLS);
  out.setColumnWidth(1, 200);
  out.setColumnWidth(2, 120);
  out.setColumnWidth(6, 160);
  out.setColumnWidth(8, 170);
  out.setColumnWidth(9, 200);
  out.setFrozenRows(4);

  ss.setActiveSheet(out);

  let mensajeFinal = `✅ INFORME BIENES PERSONALES ${anoFiscal} generado.\n\n`;
  mensajeFinal += `📊 Total Gravado:      $${totalGravado.toLocaleString('es-AR', { minimumFractionDigits: 2 })}\n`;
  mensajeFinal += `📉 MNI Aplicado:        ${mniDefinido ? '$' + mniAno.toLocaleString('es-AR', { minimumFractionDigits: 2 }) : 'NO DEFINIDO'}\n`;
  mensajeFinal += `📐 Base Imponible:     $${baseImponible.toLocaleString('es-AR', { minimumFractionDigits: 2 })}\n`;
  mensajeFinal += `💸 IMPUESTO ESTIMADO: ${(avisoMniPendiente || avisoEscalaPendiente) ? 'N/D (faltan parámetros)' : '$' + impuesto.toLocaleString('es-AR', { minimumFractionDigits: 2 })}\n`;
  if (hayPreciosFaltantes.length > 0) {
    mensajeFinal += `\n⚠️ ${hayPreciosFaltantes.length} ticker(s) sin precio: ${hayPreciosFaltantes.join(', ')}\nRevisarlos manualmente.`;
  }
  if (avisoMniPendiente || avisoEscalaPendiente) {
    mensajeFinal += `\n\n⚠️ Completá PARAM_FISCAL para ${anoFiscal} y volvé a generar el informe.`;
  }
  ui.alert(mensajeFinal);
}