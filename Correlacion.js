// =================================================================================
// ===   TITANIUM v2 — CORRELACION.GS                                          ===
// ===   Análisis de correlación entre Cedears + detección de solapamientos    ===
// ===   Sin onOpen() propio — vive unificado en Triggers.gs                   ===
// =================================================================================

var TICKER_MAP = {
  'BA.C': 'BAC'
};

function limpiarTemporales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var count = 0;
  ss.getSheets().forEach(function(s) {
    if (s.getName().indexOf('TEMP_') === 0) { ss.deleteSheet(s); count++; }
  });
  SpreadsheetApp.getUi().alert('✅ Eliminadas ' + count + ' hojas temporales.');
}

// ------------------------------------------------
function leerCartera() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Cartera');
  if (!sheet) throw new Error('No se encontró la hoja Cartera');
  var data    = sheet.getDataRange().getValues();
  var cedears = [];
  for (var i = 1; i < data.length; i++) {
    var tipo   = String(data[i][1] || '').trim();
    var ticker = String(data[i][0] || '').trim().toUpperCase();
    if (tipo === 'Cedear' && ticker) {
      cedears.push({
        ticker:   ticker,
        valorUSD: parseFloat(String(data[i][5]).replace(/[$,]/g, '')) || 0, // col F
        pl:       data[i][7] // col H — P/L % (corregido, antes apuntaba a col I)
      });
    }
  }
  return cedears;
}

// ------------------------------------------------
function generarAnalisisCompleto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var tz = Session.getScriptTimeZone();

  var cartera;
  try { cartera = leerCartera(); }
  catch (e) { ui.alert('❌ Error: ' + e.message); return; }
  if (!cartera.length) { ui.alert('❌ No se encontraron CEDEARs.'); return; }

  var tickers = cartera.map(function(c) { return c.ticker; });
  var n       = tickers.length;

  var hoy         = new Date();
  var fechaFin    = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  var fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  fechaInicio.setDate(fechaInicio.getDate() - 365);

  function fmtDisplay(d) { return Utilities.formatDate(d, tz, 'dd/MM/yyyy'); }
  function fmtKey(d)     { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  function dateGF(d)     {
    return 'DATE(' + d.getFullYear() + ',' + (d.getMonth()+1) + ',' + d.getDate() + ')';
  }

  var corrSheet = ss.getSheetByName('Correlacion_Cedears');
  if (!corrSheet) corrSheet = ss.insertSheet('Correlacion_Cedears');
  else corrSheet.clear();

  var colsNecesarias = n + 2;
  if (corrSheet.getMaxColumns() < colsNecesarias) {
    corrSheet.insertColumnsAfter(corrSheet.getMaxColumns(), colsNecesarias - corrSheet.getMaxColumns());
  }

  corrSheet.getRange('A1').setValue('⏳ Iniciando...')
    .setFontSize(12).setFontWeight('bold');
  SpreadsheetApp.flush();

  var retornosPorFecha = {};

  for (var idx = 0; idx < n; idx++) {
    var ticker   = tickers[idx];
    var gfTicker = TICKER_MAP[ticker] || ticker;

    corrSheet.getRange('A1').setValue('⏳ ' + (idx+1) + '/' + n + ': ' + ticker);
    SpreadsheetApp.flush();

    try {
      var tempName  = 'TEMP_' + ticker;
      var tempSheet = ss.getSheetByName(tempName);
      if (tempSheet) tempSheet.clear();
      else tempSheet = ss.insertSheet(tempName);

      tempSheet.getRange('A1').setFormula(
        '=GOOGLEFINANCE("' + gfTicker + '","close",' +
        dateGF(fechaInicio) + ',' + dateGF(fechaFin) + ',"DAILY")'
      );

      SpreadsheetApp.flush();
      Utilities.sleep(2500);

      var values = tempSheet.getDataRange().getValues();
      ss.deleteSheet(tempSheet);

      if (values.length < 30) {
        Logger.log('Pocos datos para ' + ticker + ': ' + values.length + ' filas');
        continue;
      }

      var precios = {};
      for (var r = 1; r < values.length; r++) {
        var fecha  = values[r][0];
        var precio = values[r][1];
        if (fecha instanceof Date && precio > 0) {
          precios[fmtKey(fecha)] = precio;
        }
      }

      var fechasOrd = Object.keys(precios).sort();
      var retMap    = {};
      for (var f = 1; f < fechasOrd.length; f++) {
        retMap[fechasOrd[f]] = (precios[fechasOrd[f]] / precios[fechasOrd[f-1]]) - 1;
      }

      if (Object.keys(retMap).length > 15) {
        retornosPorFecha[ticker] = retMap;
        Logger.log('OK: ' + ticker + ' — ' + Object.keys(retMap).length + ' días');
      }

    } catch (e) {
      Logger.log('Error ' + ticker + ': ' + e.message);
      try {
        var ts = ss.getSheetByName('TEMP_' + ticker);
        if (ts) ss.deleteSheet(ts);
      } catch (_) {}
    }
  }

  var tickersConDatos = Object.keys(retornosPorFecha).length;

  var matrizCorr = {};
  for (var i = 0; i < n; i++) {
    for (var j = i + 1; j < n; j++) {
      var tA = tickers[i], tB = tickers[j];
      var mA = retornosPorFecha[tA], mB = retornosPorFecha[tB];
      if (!mA || !mB) continue;

      var comunes = Object.keys(mA).filter(function(f) {
        return mB[f] !== undefined;
      }).sort();
      if (comunes.length < 8) continue;

      var arr1 = comunes.map(function(f) { return mA[f]; });
      var arr2 = comunes.map(function(f) { return mB[f]; });
      var key  = [tA, tB].sort().join('|');
      matrizCorr[key] = calcularCorrelacion(arr1, arr2);
    }
  }

  function getCorr(tA, tB) {
    var key = [tA, tB].sort().join('|');
    return matrizCorr.hasOwnProperty(key) ? matrizCorr[key] : null;
  }

  corrSheet.getRange('A1').setValue('📊 MATRIZ DE CORRELACIÓN CEDEARS')
    .setFontSize(14).setFontWeight('bold');
  corrSheet.getRange('A3').setValue('Período:');
  corrSheet.getRange('B3').setValue(fmtDisplay(fechaInicio) + ' → ' + fmtDisplay(fechaFin));
  corrSheet.getRange('A4').setValue('Tickers con datos:');
  corrSheet.getRange('B4').setValue(tickersConDatos + ' / ' + n);

  var ROW_H = 5;
  corrSheet.getRange(ROW_H, 1).setValue('Ticker').setFontWeight('bold');
  for (var ti = 0; ti < n; ti++) {
    corrSheet.getRange(ROW_H, ti + 2)
      .setValue(tickers[ti]).setFontWeight('bold').setHorizontalAlignment('center');
  }

  for (var i = 0; i < n; i++) {
    var rowIdx = ROW_H + 1 + i;
    corrSheet.getRange(rowIdx, 1).setValue(tickers[i]).setFontWeight('bold');
    for (var j = 0; j < n; j++) {
      var cell = corrSheet.getRange(rowIdx, j + 2);
      if (i === j) {
        cell.setValue(1).setBackground('#d9ead3').setHorizontalAlignment('center');
        continue;
      }
      var corr = getCorr(tickers[i], tickers[j]);
      if (corr === null) {
        cell.setValue('N/D').setFontColor('#999999').setHorizontalAlignment('center');
      } else {
        cell.setValue(corr).setNumberFormat('0.00').setHorizontalAlignment('center');
        if      (corr > 0.85) cell.setBackground('#f4cccc');
        else if (corr > 0.70) cell.setBackground('#fce5cd');
        else if (corr > 0.50) cell.setBackground('#fff2cc');
        else if (corr < 0.10) cell.setBackground('#d0e0e3');
      }
    }
  }

  corrSheet.autoResizeColumns(1, n + 2);
  corrSheet.getRange(ROW_H, 1, n + 1, n + 1)
    .setBorder(true, true, true, true, true, true);

  generarAvisos(ss, cartera, matrizCorr);
  var avisosSheet = ss.getSheetByName('Avisos_Cedears');
  if (avisosSheet) ss.setActiveSheet(avisosSheet);

  ui.alert(
    '✅ ¡Análisis completo!\n\n' +
    '• ' + tickersConDatos + ' / ' + n + ' tickers con datos\n' +
    '• ' + fmtDisplay(fechaInicio) + ' → ' + fmtDisplay(fechaFin) + '\n\n' +
    'Revisá la hoja "Avisos_Cedears".'
  );
}

// ------------------------------------------------
function generarAvisos(ss, cartera, matrizCorr) {
  var sheet = ss.getSheetByName('Avisos_Cedears');
  if (!sheet) sheet = ss.insertSheet('Avisos_Cedears');
  else sheet.clear();

  var carteraMap = {};
  cartera.forEach(function(c) { carteraMap[c.ticker] = c; });
  var totalUSD = cartera.reduce(function(s, c) { return s + c.valorUSD; }, 0);
  var row = 1;

  function sectionHeader(txt, color) {
    sheet.getRange(row, 1, 1, 6).merge().setBackground(color);
    sheet.getRange(row, 1).setValue(txt).setFontWeight('bold').setFontSize(12);
    row++;
  }
  function tableHeader(cols) {
    cols.forEach(function(c, i) {
      sheet.getRange(row, i+1).setValue(c).setFontWeight('bold').setBackground('#efefef');
    });
    row++;
  }
  function writePL(cell, val) {
    if (val === null || val === undefined || val === '') return;
    var num = parseFloat(String(val).replace(/[%,\s]/g, ''));
    if (isNaN(num)) { cell.setValue(val); return; }
    var frac = (Math.abs(num) > 1) ? num / 100 : num;
    cell.setValue(frac).setNumberFormat('0.0%')
        .setFontColor(frac < 0 ? '#cc0000' : '#2d6a2d');
  }

  // ═══ 1 — RESUMEN (SIN "% especie", eliminada) ═══════════════════
  sectionHeader('📋  RESUMEN CEDEARs  (por valor USD)', '#c9daf8');
  tableHeader(['Ticker', 'Valor USD', 'P/L %', '% s/total CEDEARs']);

  cartera.slice().sort(function(a,b){ return b.valorUSD - a.valorUSD; })
    .forEach(function(c) {
      sheet.getRange(row,1).setValue(c.ticker);
      sheet.getRange(row,2).setValue(c.valorUSD).setNumberFormat('"$"#,##0');
      writePL(sheet.getRange(row,3), c.pl);
      sheet.getRange(row,4).setValue(totalUSD>0 ? c.valorUSD/totalUSD : 0).setNumberFormat('0.00%');
      row++;
    });
  sheet.getRange(row,1).setValue('TOTAL').setFontWeight('bold');
  sheet.getRange(row,2).setValue(totalUSD).setNumberFormat('"$"#,##0').setFontWeight('bold');
  row += 2;

  // ═══ 2 — ETFs SOLAPADOS ════════════════════════
  sectionHeader('⚠️  ETFs SOLAPADOS — replican índices americanos similares', '#fff2cc');
  tableHeader(['Ticker', 'Valor USD', '% s/total CEDEARs', 'P/L %', 'Índice que replica']);

  var etfDesc = {
    'SPY': 'S&P 500 (cap. ponderada)',
    'QQQ': 'Nasdaq-100',
    'RSP': 'S&P 500 (igual ponderación)',
    'DIA': 'Dow Jones Industrial'
  };
  var totalEtf = 0;
  ['SPY','QQQ','RSP','DIA'].forEach(function(t) {
    var c = carteraMap[t];
    if (!c) return;
    sheet.getRange(row,1).setValue(t);
    sheet.getRange(row,2).setValue(c.valorUSD).setNumberFormat('"$"#,##0');
    sheet.getRange(row,3).setValue(totalUSD>0 ? c.valorUSD/totalUSD : 0).setNumberFormat('0.00%');
    writePL(sheet.getRange(row,4), c.pl);
    sheet.getRange(row,5).setValue(etfDesc[t]);
    totalEtf += c.valorUSD;
    row++;
  });
  sheet.getRange(row,1).setValue('TOTAL ETFs').setFontWeight('bold');
  sheet.getRange(row,2).setValue(totalEtf).setNumberFormat('"$"#,##0').setFontWeight('bold');
  sheet.getRange(row,3).setValue(totalUSD>0 ? totalEtf/totalUSD : 0).setNumberFormat('0.00%').setFontWeight('bold');
  row += 2;

  // ═══ 3 — ALTA CORRELACIÓN > 0.85 ══════════════
  sectionHeader('🔴  PARES CON ALTA CORRELACIÓN  (> 0.85)', '#f4cccc');
  tableHeader(['Ticker A', 'P/L A', 'Ticker B', 'P/L B', 'Correlación', 'Menor P/L']);

  var altos = Object.keys(matrizCorr)
    .filter(function(k){ return matrizCorr[k] > 0.85; })
    .sort(function(a,b){ return matrizCorr[b]-matrizCorr[a]; });

  if (!altos.length) {
    sheet.getRange(row,1).setValue('✅ No hay pares con correlación > 0.85'); row++;
  } else {
    altos.forEach(function(key) {
      var p=key.split('|'), tA=p[0], tB=p[1];
      var cA=carteraMap[tA], cB=carteraMap[tB];
      sheet.getRange(row,1).setValue(tA);
      if (cA) writePL(sheet.getRange(row,2), cA.pl);
      sheet.getRange(row,3).setValue(tB);
      if (cB) writePL(sheet.getRange(row,4), cB.pl);
      sheet.getRange(row,5).setValue(matrizCorr[key]).setNumberFormat('0.00').setBackground('#f4cccc');
      if (cA && cB) {
        var pA=parseFloat(String(cA.pl).replace(/[%,\s]/g,''));
        var pB=parseFloat(String(cB.pl).replace(/[%,\s]/g,''));
        if (!isNaN(pA) && !isNaN(pB)) {
          sheet.getRange(row,6).setValue(pA<=pB?tA:tB)
            .setFontWeight('bold').setFontColor('#cc0000');
        }
      }
      row++;
    });
  }
  row++;

  // ═══ 4 — MODERADA 0.70-0.85 ════════════════════
  sectionHeader('🟡  CORRELACIÓN MODERADA  (0.70 – 0.85)', '#fce5cd');
  tableHeader(['Ticker A', 'P/L A', 'Ticker B', 'P/L B', 'Correlación', '']);

  var mods = Object.keys(matrizCorr)
    .filter(function(k){ return matrizCorr[k]>0.70 && matrizCorr[k]<=0.85; })
    .sort(function(a,b){ return matrizCorr[b]-matrizCorr[a]; });

  if (!mods.length) {
    sheet.getRange(row,1).setValue('✅ No hay pares con correlación 0.70–0.85'); row++;
  } else {
    mods.forEach(function(key) {
      var p=key.split('|'), tA=p[0], tB=p[1];
      var cA=carteraMap[tA], cB=carteraMap[tB];
      sheet.getRange(row,1).setValue(tA);
      if (cA) writePL(sheet.getRange(row,2), cA.pl);
      sheet.getRange(row,3).setValue(tB);
      if (cB) writePL(sheet.getRange(row,4), cB.pl);
      sheet.getRange(row,5).setValue(matrizCorr[key]).setNumberFormat('0.00').setBackground('#fce5cd');
      row++;
    });
  }
  row++;

  // ═══ 5 — BAJA CORRELACIÓN < 0.10 ═══
  sectionHeader('🟢  BAJA CORRELACIÓN  (< 0.10) — buenos diversificadores', '#d9ead3');
  tableHeader(['Ticker A', 'P/L A', 'Ticker B', 'P/L B', 'Correlación', '']);

  var bajos = Object.keys(matrizCorr)
    .filter(function(k){ return matrizCorr[k] < 0.10; })
    .sort(function(a,b){ return matrizCorr[a]-matrizCorr[b]; })
    .slice(0, 15);

  if (!bajos.length) {
    sheet.getRange(row,1).setValue('No hay pares con correlación < 0.10'); row++;
  } else {
    bajos.forEach(function(key) {
      var p=key.split('|'), tA=p[0], tB=p[1];
      var cA=carteraMap[tA], cB=carteraMap[tB];
      sheet.getRange(row,1).setValue(tA);
      if (cA) writePL(sheet.getRange(row,2), cA.pl);
      sheet.getRange(row,3).setValue(tB);
      if (cB) writePL(sheet.getRange(row,4), cB.pl);
      sheet.getRange(row,5).setValue(matrizCorr[key]).setNumberFormat('0.00').setBackground('#d9ead3');
      row++;
    });
  }

  sheet.autoResizeColumns(1, 6);
}

// ------------------------------------------------
function calcularCorrelacion(arr1, arr2) {
  var n=arr1.length;
  if (n<8) return 0;
  var s1=0,s2=0,s12=0,s1q=0,s2q=0;
  for (var i=0;i<n;i++){
    var x=arr1[i],y=arr2[i];
    s1+=x;s2+=y;s12+=x*y;s1q+=x*x;s2q+=y*y;
  }
  var num=n*s12-s1*s2;
  var den=Math.sqrt((n*s1q-s1*s1)*(n*s2q-s2*s2));
  return den===0?0:num/den;
}
