// =================================================================================
// ===   TITANIUM v2 — AUDITORIA.GS                                             ===
// ===   Genera AUDITORIA_DETALLE y AUDITORIA_RESUMEN                          ===
// ===   Reutiliza calcXIRR(), cleanNum(), FECHA_CORTE_CAJA, SALDO_INICIAL_CAJA ===
// ===   de Motor.gs — sin duplicar lógica.                                     ===
// =================================================================================

const SHEETS_AUDIT = {
  DETALLE: "AUDITORIA_DETALLE",
  RESUMEN: "AUDITORIA_RESUMEN"
};

function GENERAR_AUDITORIA_COMPLETA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. CCL — ahora vive en Precios!B4 (ya no hay hoja CCL separada)
  let ccl = 1000;
  try {
    const val = ss.getSheetByName(HOJAS.PRECIOS).getRange("B4").getValue();
    if (typeof val === 'number' && val > 500) ccl = val;
  } catch (e) {
    Logger.log("Error al leer CCL, usando $1000 por defecto: " + e.message);
  }

  // 2. Precios: merge de flags USD/PESOS igual que Motor.gs
  const precios = {};
  try {
    const preciosSheet = ss.getSheetByName(HOJAS.PRECIOS);
    if (preciosSheet) {
      const data = preciosSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const tk = String(data[i][0]).toUpperCase().trim();
        if (!tk) continue;

        const pr = cleanNum(data[i][1]);
        const moneda = String(data[i][2]).toUpperCase().trim();

        if (pr === 0 && !moneda) continue;
        if (!precios[tk]) precios[tk] = { price: 0, enUSD: false, enPesos: false };
        if (pr > 0) precios[tk].price = pr;

        if (moneda === 'USD')   { precios[tk].enUSD = true;  precios[tk].enPesos = false; }
        if (moneda === 'PESOS') { precios[tk].enPesos = true; precios[tk].enUSD = false; }
      }
    }
  } catch (e) {
    Logger.log("Error al leer precios actuales: " + e.message);
  }

  // 3. Log de transacciones
  const logSheet = ss.getSheetByName(HOJAS.LOG);
  if (!logSheet) {
    ui.alert("❌ Error: No se encontró la hoja de transacciones '" + HOJAS.LOG + "'.");
    return;
  }

  const logData = logSheet.getDataRange().getValues();
  const logSinHeader = logData.slice(1);
  logSinHeader.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  // 4. Caja virtual — reutiliza constantes de Motor.gs
  let cajaVirtual = SALDO_INICIAL_CAJA;

  logSinHeader.forEach(row => {
    const fecha = new Date(row[1]);
    if (isNaN(fecha.getTime()) || fecha < FECHA_CORTE_CAJA) return;

    const ticker = String(row[3]).toUpperCase().trim();
    const mov = String(row[5]).toLowerCase().trim();
    const montoUSD = Math.abs(cleanNum(row[10]));

    if (ticker === 'USD' || ticker === 'CASH') {
      if (mov.includes('aporte') || mov.includes('compra') || mov.includes('suscripcion') || mov.includes('canje_entrada')) {
        cajaVirtual += montoUSD;
      } else if (mov.includes('retiro') || mov.includes('venta') || mov.includes('rescate') || mov.includes('canje_salida')) {
        cajaVirtual -= montoUSD;
      }
    } else {
      if (mov.includes('compra') || mov.includes('suscripcion') || mov.includes('canje_entrada')) {
        cajaVirtual -= montoUSD;
      } else if (mov.includes('venta') || mov.includes('rescate') || mov.includes('canje_salida')) {
        cajaVirtual += montoUSD;
      } else if (mov.includes('dividendo') || mov.includes('renta') || mov.includes('interes') || mov.includes('amortiza')) {
        cajaVirtual += montoUSD;
      }
    }
  });

  // =============================================================================
  // PASO 3.5 de la consolidación (ver debate "Portafolio Ultimate"): esta
  // función ya NO reimplementa el loop de posiciones. La posición final de
  // cada ticker sale de snapshotAFecha() y el detalle transacción por
  // transacción sale de ledgerCompleto() — ambas en MotorPosiciones.gs, que
  // ya aplican Método B, el tope de venta (FIX M-011) y la regla de
  // venta-sin-stock (FIX M-002). Se preserva la regla de negocio PROPIA de
  // esta auditoría (distinta de la del Dashboard): la ganancia extra de una
  // amortización se suma SIEMPRE a rentaRF, sin el filtro de
  // tickersConRentaExplicita — acá interesa el efectivo realmente cobrado,
  // no evitar duplicar con el Modified Dietz (que esta hoja no calcula).
  // =============================================================================
  const snapshotHoy = snapshotAFecha(logSinHeader, HOY_SIMULADA);
  const ledger = ledgerCompleto(logSinHeader);

  const portfolio = {};
  Object.keys(snapshotHoy.posiciones).forEach(tk => {
    const p = snapshotHoy.posiciones[tk];
    portfolio[tk] = {
      ticker: tk, tipo: p.tipo, qty: p.q, costo: p.costo, costoOriginal: p.costoOriginal,
      cobradoRentas: 0, cobradoAmortiz: 0, realizedPL: 0, cashflows: []
    };
  });

  const transDetalle = [];
  let totalGciaCapitalRV = 0;
  let totalGciaCapitalRF = 0;
  let totalDivsRV = 0;
  let totalRentasRF = 0;
  let totalAmortizRF = 0;
  const gciaPorAnio = {};

  let qAntesPorTicker = {};
  let costoAntesPorTicker = {};

  ledger.transacciones.forEach(t => {
    const year = t.anio;
    const esRV = t.tipo.toLowerCase().includes('cedear') || t.tipo.toLowerCase().includes('accion');
    const qAntes = qAntesPorTicker[t.ticker] || 0;
    const costoAntes = (costoAntesPorTicker[t.ticker] !== undefined) ? costoAntesPorTicker[t.ticker] : 0;
    const costoUsadoEstaFila = costoAntes - t.costoDespues; // > 0 en ventas/amortizaciones que reducen costo

    if (!portfolio[t.ticker]) {
      // Ticker sin posición viva hoy (se cerró del todo en el pasado) pero con
      // historia — lo agregamos igual para que su detalle no se pierda.
      portfolio[t.ticker] = {
        ticker: t.ticker, tipo: t.tipo, qty: 0, costo: 0, costoOriginal: 0,
        cobradoRentas: 0, cobradoAmortiz: 0, realizedPL: 0, cashflows: []
      };
    }
    const p = portfolio[t.ticker];

    if (t.cashflow) p.cashflows.push({ date: t.fecha, amount: t.cashflow });

    if (!gciaPorAnio[year]) gciaPorAnio[year] = { capRV: 0, capRF: 0, divRV: 0, rentaRF: 0, amortizRF: 0 };

    if (t.movimiento.includes('venta') || t.movimiento.includes('rescate') || t.movimiento.includes('canje_salida')) {
      const etiqueta = (qAntes > 0) ? "VENTA" : "VENTA (sin compra previa)";
      p.realizedPL += t.gananciaRealizada;
      transDetalle.push([t.fecha, t.ticker, t.tipo, etiqueta, t.montoUSD, costoUsadoEstaFila, t.gananciaRealizada, year]);

      if (esRV) { totalGciaCapitalRV += t.gananciaRealizada; gciaPorAnio[year].capRV += t.gananciaRealizada; }
      else { totalGciaCapitalRF += t.gananciaRealizada; gciaPorAnio[year].capRF += t.gananciaRealizada; }

    } else if (t.movimiento.includes('dividendo') || t.movimiento.includes('renta') || t.movimiento.includes('interes')) {
      p.cobradoRentas += t.montoUSD;
      transDetalle.push([t.fecha, t.ticker, t.tipo, esRV ? "DIVIDENDO" : "RENTA", t.montoUSD, 0, t.montoUSD, year]);

      if (esRV) { totalDivsRV += t.montoUSD; gciaPorAnio[year].divRV += t.montoUSD; }
      else { totalRentasRF += t.montoUSD; gciaPorAnio[year].rentaRF += t.montoUSD; }

    } else if (t.movimiento.includes('amortiza')) {
      // Nota: filas de amortización con Monto_Neto_USD = 0 nunca llegan hasta
      // acá — ledgerCompleto() ya las filtra (con su warning correspondiente)
      // antes de incluirlas en la lista de transacciones.
      p.cobradoAmortiz += t.montoUSD;
      transDetalle.push([t.fecha, t.ticker, t.tipo, "AMORTIZACION", t.montoUSD, costoUsadoEstaFila, t.gananciaRealizada, year]);

      totalAmortizRF += t.montoUSD;
      gciaPorAnio[year].amortizRF += t.montoUSD;

      if (t.gananciaRealizada > 0) {
        gciaPorAnio[year].rentaRF += t.gananciaRealizada;
        totalRentasRF += t.gananciaRealizada;
      }
    }
    // Nota: 'compra' y 'split' no generan fila en transDetalle — igual que antes.

    qAntesPorTicker[t.ticker] = t.qDespues;
    costoAntesPorTicker[t.ticker] = t.costoDespues;
  });

  // 7. Valuación final + XIRR (usa calcXIRR de Motor.gs)
  const tickersResumen = [];
  let totalValuationUSD = 0, totalCostoResidualUSD = 0, totalCostoOriginalUSD = 0;
  let totalCobradoDivRentaUSD = 0, totalCobradoAmortizUSD = 0, totalGciaRealizadaGlob = 0;

  Object.keys(portfolio).forEach(tk => {
    const p = portfolio[tk];
    if (p.qty <= 0 && p.realizedPL === 0 && p.cobradoRentas === 0 && p.cobradoAmortiz === 0) return;

    const precioEntry = precios[tk] || { price: 0, enUSD: false, enPesos: false };
    const precioRaw = precioEntry.price;
    const enUSD = precioEntry.enUSD;
    const enPesos = precioEntry.enPesos;
    let valMercadoUSD = 0, precioRefUSD = 0;

    if (enUSD) {
      precioRefUSD = precioRaw;
      valMercadoUSD = p.qty * precioRefUSD;
    } else if (enPesos) {
      precioRefUSD = precioRaw / ccl;
      valMercadoUSD = p.qty * precioRefUSD;
    } else {
      const esBonoON = p.tipo.toLowerCase().includes('bono') || p.tipo.toLowerCase().includes('negociable') || p.tipo.toLowerCase().includes('renta fija');
      if (esBonoON) {
        precioRefUSD = (precioRaw / ccl) / 100;
      } else {
        precioRefUSD = precioRaw / ccl;
      }
      valMercadoUSD = p.qty * precioRefUSD;
    }

    let gciaNoRealizada = 0;
    if (p.qty > 0) gciaNoRealizada = valMercadoUSD - p.costo;

    const cfsXirr = [...p.cashflows];
    if (p.qty > 0 && valMercadoUSD > 0) cfsXirr.push({ date: new Date(), amount: valMercadoUSD });

    // Reutiliza calcXIRR() de Motor.gs — necesita { d, v }, no { date, amount }
    const cfsAdaptado = cfsXirr.map(c => ({ d: c.date, v: c.amount }));
    const xirrCalculado = calcXIRR(cfsAdaptado);

    tickersResumen.push({
      ticker: tk, tipo: p.tipo, qty: p.qty, costo: p.costo, costoOriginal: p.costoOriginal,
      valMercado: valMercadoUSD, gciaNoRealizada: gciaNoRealizada, gciaRealizada: p.realizedPL,
      cobradoRentas: p.cobradoRentas, cobradoAmortiz: p.cobradoAmortiz, xirr: xirrCalculado
    });

    totalValuationUSD += valMercadoUSD;
    totalCostoResidualUSD += p.costo;
    totalCostoOriginalUSD += p.costoOriginal;
    totalCobradoDivRentaUSD += p.cobradoRentas;
    totalCobradoAmortizUSD += p.cobradoAmortiz;
    totalGciaRealizadaGlob += p.realizedPL;
  });

  // 8. Fila de liquidez
  tickersResumen.push({
    ticker: "USD", tipo: "Liquidez", qty: cajaVirtual, costo: cajaVirtual, costoOriginal: cajaVirtual,
    valMercado: cajaVirtual, gciaNoRealizada: 0, gciaRealizada: 0, cobradoRentas: 0, cobradoAmortiz: 0, xirr: 0
  });
  totalValuationUSD += cajaVirtual;
  totalCostoResidualUSD += cajaVirtual;
  totalCostoOriginalUSD += cajaVirtual;

  tickersResumen.sort((a, b) => b.valMercado - a.valMercado);

  // 9. HOJA 1: AUDITORIA_DETALLE
  let sheetDetalle = ss.getSheetByName(SHEETS_AUDIT.DETALLE);
  if (!sheetDetalle) sheetDetalle = ss.insertSheet(SHEETS_AUDIT.DETALLE);
  sheetDetalle.clear();
  sheetDetalle.clearFormats();

  sheetDetalle.getRange(1, 1, 1, 8).setValues([["FECHA", "TICKER", "TIPO ACTIVO", "OPERACIÓN", "MONTO COBRADO/VENTA (USD)", "COSTO HISTÓRICO (USD)", "GANANCIA REALIZADA (USD)", "AÑO"]])
    .setFontWeight("bold").setBackground("#0a192f").setFontColor("#FFC300").setHorizontalAlignment("center");

  transDetalle.sort((a, b) => a[0] - b[0]);

  if (transDetalle.length > 0) {
    sheetDetalle.getRange(2, 1, transDetalle.length, 8).setValues(transDetalle);
    sheetDetalle.getRange(2, 1, transDetalle.length, 1).setNumberFormat("dd/MM/yyyy");
    sheetDetalle.getRange(2, 5, transDetalle.length, 3).setNumberFormat("$#,##0.00");
    sheetDetalle.getRange(2, 8, transDetalle.length, 1).setNumberFormat("0");

    const rulePos = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.01).setBackground("#e8f5e9").setFontColor("#2e7d32")
      .setRanges([sheetDetalle.getRange(2, 7, transDetalle.length, 1)]).build();
    const ruleNeg = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(-0.01).setBackground("#ffebee").setFontColor("#c62828")
      .setRanges([sheetDetalle.getRange(2, 7, transDetalle.length, 1)]).build();
    sheetDetalle.setConditionalFormatRules([rulePos, ruleNeg]);
  }
  sheetDetalle.autoResizeColumns(1, 8);

  // 10. HOJA 2: AUDITORIA_RESUMEN
  let sheetResumen = ss.getSheetByName(SHEETS_AUDIT.RESUMEN);
  if (!sheetResumen) sheetResumen = ss.insertSheet(SHEETS_AUDIT.RESUMEN);
  sheetResumen.clear();
  sheetResumen.clearFormats();

  let rowIdx = 1;

  sheetResumen.getRange(rowIdx, 1, 1, 10).merge()
    .setValue("AUDITORÍA DE PORTAFOLIO - RESUMEN GENERAL (VALORES EN USD)")
    .setFontWeight("bold").setFontSize(13).setBackground("#0a192f").setFontColor("#FFC300").setHorizontalAlignment("center");
  sheetResumen.setRowHeight(rowIdx, 30);
  rowIdx += 2;

  const totalGciaNoReal = totalValuationUSD - totalCostoResidualUSD;
  const totalPLHistorico = totalGciaRealizadaGlob + totalGciaNoReal + totalCobradoDivRentaUSD;

  const kpis = [
    ["Valor Total Mercado Hoy:", totalValuationUSD, "Costo Original Total:", totalCostoOriginalUSD],
    ["Ganancia No Realizada (Latente):", totalGciaNoReal, "Costo Residual Libros:", totalCostoResidualUSD],
    ["Ganancia Realizada (Ventas):", totalGciaRealizadaGlob, "Total Dividendos/Rentas Cobrados:", totalCobradoDivRentaUSD],
    ["Total Amortizaciones Cobradas:", totalCobradoAmortizUSD, "P&L TOTAL HISTÓRICO CONSOLIDADO:", totalPLHistorico]
  ];

  sheetResumen.getRange(rowIdx, 1, 4, 4).setValues(kpis);
  sheetResumen.getRange(rowIdx, 1, 4, 1).setFontWeight("bold");
  sheetResumen.getRange(rowIdx, 3, 4, 1).setFontWeight("bold");
  sheetResumen.getRange(rowIdx, 2, 4, 1).setNumberFormat("$#,##0.00");
  sheetResumen.getRange(rowIdx, 4, 4, 1).setNumberFormat("$#,##0.00");
  sheetResumen.getRange(rowIdx + 3, 3, 1, 2).setFontWeight("bold").setBackground("#fff8e1").setFontColor("#b78103");

  rowIdx += 6;

  sheetResumen.getRange(rowIdx, 1, 1, 10).merge()
    .setValue("AUDITORÍA DETALLADA POR ACTIVO (POSICIONES ABIERTAS Y CERRADAS)")
    .setFontWeight("bold").setFontSize(11).setBackground("#1a3a5c").setFontColor("white").setHorizontalAlignment("center");
  rowIdx++;

  const headersB = ["TICKER", "TIPO ACTIVO", "CANT. NOMINAL", "COSTO RESID (USD)", "COSTO ORIG (USD)", "VALOR MERCADO (USD)", "GCIA NO REAL (USD)", "GCIA REALIZ (USD)", "COBRADO DIVS/RENT (USD)", "XIRR %"];
  sheetResumen.getRange(rowIdx, 1, 1, 10).setValues([headersB])
    .setFontWeight("bold").setBackground("#f0f4f8").setHorizontalAlignment("center");
  rowIdx++;

  const startRowB = rowIdx;
  tickersResumen.forEach(r => {
    sheetResumen.getRange(rowIdx, 1, 1, 10).setValues([[
      r.ticker, r.tipo, r.qty, r.costo, r.costoOriginal, r.valMercado, r.gciaNoRealizada, r.gciaRealizada, r.cobradoRentas, r.xirr
    ]]);

    sheetResumen.getRange(rowIdx, 3).setNumberFormat("#,##0.0000");
    sheetResumen.getRange(rowIdx, 4, 1, 6).setNumberFormat("$#,##0.00");
    sheetResumen.getRange(rowIdx, 10).setNumberFormat("0.00%");

    if (r.qty < 0.0001) {
      sheetResumen.getRange(rowIdx, 1, 1, 10).setFontColor("#888888");
      sheetResumen.getRange(rowIdx, 3).setValue("-");
      sheetResumen.getRange(rowIdx, 6).setValue("-");
    }
    if (r.ticker === 'USD') sheetResumen.getRange(rowIdx, 10).setValue("-");

    rowIdx++;
  });

  if (tickersResumen.length > 0) {
    const rangeXirr = sheetResumen.getRange(startRowB, 10, tickersResumen.length, 1);
    const ruleXirrPos = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.0001)
      .setFontColor("#2e7d32").setBold(true).setRanges([rangeXirr]).build();
    const ruleXirrNeg = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(-0.0001)
      .setFontColor("#c62828").setBold(true).setRanges([rangeXirr]).build();

    const rangeNoReal = sheetResumen.getRange(startRowB, 7, tickersResumen.length, 1);
    const ruleNoRealPos = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.01)
      .setBackground("#e8f5e9").setFontColor("#2e7d32").setRanges([rangeNoReal]).build();
    const ruleNoRealNeg = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(-0.01)
      .setBackground("#ffebee").setFontColor("#c62828").setRanges([rangeNoReal]).build();

    sheetResumen.setConditionalFormatRules([ruleXirrPos, ruleXirrNeg, ruleNoRealPos, ruleNoRealNeg]);
  }

  rowIdx += 2;

  sheetResumen.getRange(rowIdx, 1, 1, 7).merge()
    .setValue("RESUMEN DE GANANCIAS Y RENTAS COBRADAS POR AÑO FISCAL")
    .setFontWeight("bold").setFontSize(11).setBackground("#1a3a5c").setFontColor("white").setHorizontalAlignment("center");
  rowIdx++;

  const headersC = [
    "AÑO FISCAL",
    "GCIA CAPITAL RV (USD)",
    "GCIA CAPITAL RF (USD)",
    "DIVIDENDOS RV (USD)",
    "RENTAS/INTERES RF (USD)",
    "AMORTIZACIONES RF (USD)",
    "TOTAL COBRADO/REAL (USD)"
  ];
  sheetResumen.getRange(rowIdx, 1, 1, 7).setValues([headersC])
    .setFontWeight("bold").setBackground("#f0f4f8").setHorizontalAlignment("center");
  rowIdx++;

  const añosSorted = Object.keys(gciaPorAnio).sort();
  const startRowC = rowIdx;

  añosSorted.forEach(y => {
    const yr = gciaPorAnio[y];
    const totalYr = (yr.capRV || 0) + (yr.capRF || 0) + (yr.divRV || 0) + (yr.rentaRF || 0) + (yr.amortizRF || 0);
    sheetResumen.getRange(rowIdx, 1, 1, 7).setValues([[
      y,
      yr.capRV || 0,
      yr.capRF || 0,
      yr.divRV || 0,
      yr.rentaRF || 0,
      yr.amortizRF || 0,
      totalYr
    ]]);
    sheetResumen.getRange(rowIdx, 1).setNumberFormat("0").setHorizontalAlignment("center");
    sheetResumen.getRange(rowIdx, 2, 1, 6).setNumberFormat("$#,##0.00");
    rowIdx++;
  });

  if (añosSorted.length > 0) {
    sheetResumen.getRange(rowIdx, 1).setValue("Total Consolidado").setFontWeight("bold").setHorizontalAlignment("right");
    sheetResumen.getRange(rowIdx, 2).setFormula(`=SUM(B${startRowC}:B${rowIdx - 1})`);
    sheetResumen.getRange(rowIdx, 3).setFormula(`=SUM(C${startRowC}:C${rowIdx - 1})`);
    sheetResumen.getRange(rowIdx, 4).setFormula(`=SUM(D${startRowC}:D${rowIdx - 1})`);
    sheetResumen.getRange(rowIdx, 5).setFormula(`=SUM(E${startRowC}:E${rowIdx - 1})`);
    sheetResumen.getRange(rowIdx, 6).setFormula(`=SUM(F${startRowC}:F${rowIdx - 1})`);
    sheetResumen.getRange(rowIdx, 7).setFormula(`=SUM(G${startRowC}:G${rowIdx - 1})`);
    sheetResumen.getRange(rowIdx, 1, 1, 7).setFontWeight("bold").setBackground("#e8f5e9");
    sheetResumen.getRange(rowIdx, 2, 1, 6).setNumberFormat("$#,##0.00");
  }

  sheetResumen.autoResizeColumns(1, 10);
  ss.setActiveSheet(sheetResumen);

  ui.alert("✅ Auditoría Generada Exitosamente.\n\nSe actualizaron las hojas:\n1. AUDITORIA_RESUMEN\n2. AUDITORIA_DETALLE");
}