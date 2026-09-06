// =================================================================================
// === DIAGNOSTICO.GS — Script standalone, de SOLO LECTURA                      ===
// === No modifica el Log ni ningún archivo de Motor/Auditoria/Fiscal/Bombonera. ===
// === Único efecto secundario: crea la hoja temporal "DIAGNOSTICO_TEMP".       ===
// =================================================================================

function DIAGNOSTICO_VENTA_MAYOR_A_COMPRA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Log_Transacciones_TITANIUM");
  if (!logSheet) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja Log_Transacciones_TITANIUM.");
    return;
  }

  const logData = logSheet.getDataRange().getValues();
  const logSinHeader = logData.slice(1);
  logSinHeader.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  const cleanNum = (v) => {
    if (typeof v === 'number') return v;
    if (!v) return 0;
    let s = String(v).trim().replace('%', '');
    if (s.includes('.') && s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.includes(',')) s = s.replace(',', '.');
    return parseFloat(s) || 0;
  };

  let saldoPorTicker = {};
  let incidentes = [];

  logSinHeader.forEach((row, idx) => {
    const fecha = row[1];
    const ticker = String(row[3]).toUpperCase().trim();
    const mov = String(row[5]).toLowerCase().trim();
    const cant = cleanNum(row[6]);
    const ratioSplit = cleanNum(row[11]);

    if (!ticker || ticker === 'USD' || ticker === 'CASH') return;
    if (!(ticker in saldoPorTicker)) saldoPorTicker[ticker] = 0;

    if (mov.includes('compra') || mov.includes('aporte') || mov.includes('suscripcion') || mov.includes('canje_entrada')) {
      saldoPorTicker[ticker] += cant;

    } else if (mov.includes('venta') || mov.includes('rescate') || mov.includes('canje_salida')) {
      const saldoAntes = saldoPorTicker[ticker];
      saldoPorTicker[ticker] -= cant;
      if (saldoPorTicker[ticker] < -0.0001) {
        incidentes.push({
          fila: idx + 2,
          fecha: fecha,
          ticker: ticker,
          movimiento: mov,
          cantidadVenta: cant,
          saldoAntes: saldoAntes,
          faltante: Math.abs(saldoPorTicker[ticker])
        });
      }

    } else if (mov.includes('split')) {
      if (ratioSplit > 0) saldoPorTicker[ticker] *= ratioSplit;
    }
    // Nota: 'amortiza' y 'dividendo'/'renta' no tocan cantidad (Método B) —
    // no participan de este chequeo, que es puramente de unidades.
  });

  let hojaSalida = ss.getSheetByName("DIAGNOSTICO_TEMP");
  if (hojaSalida) ss.deleteSheet(hojaSalida);
  hojaSalida = ss.insertSheet("DIAGNOSTICO_TEMP");

  hojaSalida.getRange(1, 1, 1, 7).setValues([[
    "Fila en Log", "Fecha", "Ticker", "Movimiento", "Cant. Vendida", "Saldo Antes de Vender", "Faltante"
  ]]).setFontWeight("bold");

  if (incidentes.length > 0) {
    const filas = incidentes.map(i => [i.fila, i.fecha, i.ticker, i.movimiento, i.cantidadVenta, i.saldoAntes, i.faltante]);
    hojaSalida.getRange(2, 1, filas.length, 7).setValues(filas);
    hojaSalida.getRange(2, 2, filas.length, 1).setNumberFormat("dd/MM/yyyy");
  }
  hojaSalida.autoResizeColumns(1, 7);

  const mensaje = incidentes.length === 0
    ? "✅ No se encontró ningún caso de venta mayor a la cantidad tenida hasta ese momento, en ningún ticker.\n\nNo hace falta revisar nada a mano."
    : `⚠️ Se encontraron ${incidentes.length} fila(s) donde la venta superó el saldo tenido hasta ese momento.\n\nRevisá la hoja "DIAGNOSTICO_TEMP" para el detalle.`;

  SpreadsheetApp.getUi().alert(mensaje);
  console.log(mensaje);
}