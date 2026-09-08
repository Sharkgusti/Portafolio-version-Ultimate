// =========================================================================
// BOMBONERA.GS — Conectado a Titanium (Portafolio version ANTROPIC)
// Reutiliza: calcXIRR() de Motor.gs, HOJAS de Motor.gs, cleanNum() de Motor.gs,
// calcularCantidadesNetas() de Motor.gs (tenencia real automática desde el Log)
// =========================================================================

function getSettlementDateBombonera(startDate) {
  let settlementDate = new Date(startDate);
  let daysToAdd = 1; // T+1
  while (daysToAdd > 0) {
    settlementDate.setDate(settlementDate.getDate() + 1);
    let dayOfWeek = settlementDate.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) daysToAdd--;
  }
  return settlementDate;
}

// -----------------------------------------------------------------------
// Determina la frecuencia de pago de una hoja de bono — misma lógica que
// consolidarFuturosFlujosRF() de Triggers.gs, para consistencia total.
// -----------------------------------------------------------------------
function obtenerFrecuenciaBono(sheet, data) {
  let divisorFrecuencia = 2;
  let frecuenciaManual = 0;
  if (data[0] && data[0].length > 12) {
    let val = data[0][12];
    if (typeof val === 'number' && val > 0) frecuenciaManual = val;
  }
  if (frecuenciaManual > 0) {
    divisorFrecuencia = frecuenciaManual;
  } else {
    const fechasPagos = data.slice(2).map(r => r[2]).filter(d => d instanceof Date && !isNaN(d));
    if (fechasPagos.length > 1) {
      fechasPagos.sort((a, b) => a - b);
      let sumaDias = 0, conteo = 0;
      for (let k = 1; k < fechasPagos.length; k++) {
        let diff = (fechasPagos[k] - fechasPagos[k - 1]) / (1000 * 60 * 60 * 24);
        if (diff > 20 && diff < 400) { sumaDias += diff; conteo++; }
      }
      if (conteo > 0) {
        let prom = sumaDias / conteo;
        if (prom >= 25 && prom <= 45) divisorFrecuencia = 12;
        else if (prom >= 80 && prom <= 100) divisorFrecuencia = 4;
        else if (prom >= 160 && prom <= 200) divisorFrecuencia = 2;
        else if (prom >= 340) divisorFrecuencia = 1;
      }
    }
  }
  return divisorFrecuencia;
}

// -----------------------------------------------------------------------
// Lee el cronograma de cupones futuros de una hoja de bono, SIEMPRE
// normalizado "por cada 100 nominales" — sirve igual para tenencia real
// o teórica (100), porque nunca multiplica por la cantidad.
//
// TENENCIA: ahora se lee AUTOMÁTICAMENTE desde el Log real (vía
// calcularCantidadesNetas(), de Motor.gs) en vez de tipearla a mano en K3.
// Si el ticker no aparece en tu cartera real (por ejemplo, un bono que solo
// estás comparando sin tenerlo), cae al K3 manual como respaldo — así el
// truco de "K3 = 100 para ver el flujo teórico" sigue funcionando igual.
// -----------------------------------------------------------------------
function leerFlujosPorCada100(nombreHoja, fechaCorte, cantidadesReales) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) return { flujos: [], tenenciaReal: 0 };

  const data = sheet.getDataRange().getValues();
  if (data.length < 3) return { flujos: [], tenenciaReal: 0 };

  // Tenencia automática: primero buscamos el ticker en el Log real.
  let tenenciaReal = 0;
  const tkNormalizado = String(nombreHoja).toUpperCase().trim();
  if (cantidadesReales && cantidadesReales[tkNormalizado] && cantidadesReales[tkNormalizado].q > 0.01) {
    tenenciaReal = cantidadesReales[tkNormalizado].q;
  } else {
    // Respaldo: K3 manual (sirve para tenencias teóricas de comparación, K3=100)
    let qty = 0;
    try { qty = sheet.getRange("K3").getValue(); } catch (e) { }
   tenenciaReal = (typeof qty === 'number' && qty > 0 && qty !== TENENCIA_TEORICA_BOMBONERA) ? qty : 0;
  }

  const divisorFrecuencia = obtenerFrecuenciaBono(sheet, data);
  const flujos = [];

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    let d = row[2];

    // FIX: si la fecha llegó como texto (mismo problema que ya resolvimos en
    // Triggers.gs para Proyeccion_Flujos_RF), la convertimos a Date real.
    if (typeof d === 'string' && d.includes('/')) {
      let partes = d.split('/');
      if (partes.length === 3) d = new Date(partes[2], partes[1] - 1, partes[0]);
    }

    if (d instanceof Date && !isNaN(d) && d >= fechaCorte) {
      // FIX: leer los porcentajes robustamente aunque vengan como texto
      // ("7,5%" en vez de 0.075), mismo patrón que Triggers.gs.
      let tA = parseFloat(String(row[3]).replace('%', '').replace(',', '.')) || 0;
      let tAm = parseFloat(String(row[5]).replace('%', '').replace(',', '.')) || 0;
      let valR = row[6];
      let tR = (valR !== "" && valR !== null) ? (parseFloat(String(valR).replace('%', '').replace(',', '.')) || 0) : 1;

      if (tA > 1) tA /= 100;
      if (tAm > 1) tAm /= 100;
      if (tR > 1) tR /= 100;

      // Siempre "por cada 100 nominales", sin importar la tenencia real
      const capitalPor100 = tAm * 100;
      const rentaPor100 = (tR * tA * 100) / divisorFrecuencia;
      const totalPor100 = rentaPor100 + capitalPor100;

      if (totalPor100 > 0.01) {
        flujos.push({
          cuponNro: row[0], fechaPago: d,
          tasaInteres: tA * 100, amortizacion: capitalPor100,
          residual: tR * 100, renta: rentaPor100, capital: capitalPor100, total: totalPor100
        });
      }
    }
  }
  return { flujos, tenenciaReal };
}

function calculateModifiedDurationBombonera(cashflows, dates, tir) {
  if (cashflows.length !== dates.length || cashflows.length <= 1 || tir === null || typeof tir !== 'number') return null;
  const settlementDate = dates[0];
  const dirtyPrice = -cashflows[0];
  if (dirtyPrice <= 0) return null;
  const frequency = 2;
  let sumPvCfTimesT = 0;
  for (let i = 1; i < cashflows.length; i++) {
    const t = (dates[i].getTime() - settlementDate.getTime()) / (1000 * 60 * 60 * 24 * 365.25);
    const pv_cf = cashflows[i] / Math.pow(1 + tir, t);
    sumPvCfTimesT += pv_cf * t;
  }
  const macaulayDuration = sumPvCfTimesT / dirtyPrice;
  const modifiedDuration = macaulayDuration / (1 + (tir / frequency));
  return isFinite(modifiedDuration) ? modifiedDuration : null;
}

// -----------------------------------------------------------------------
// MASTER DATA — reemplaza la vieja getMasterData()
// Lee: Universo (reemplaza a UNIVERSO), Precios!S:T + Universo!S:T
// (traducción pesos/dólar), Cartera (reemplaza a TENENCIA), hojas
// individuales de bono (reemplazan a BIBLIOTECA_DE_BONOS)
// -----------------------------------------------------------------------
function getMasterDataBombonera() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'bombonera_master_v1';
  const cachedData = cache.get(cacheKey);
  if (cachedData != null) return cachedData;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // NUEVO: tenencia real, leída automáticamente del Log (misma fuente de
    // verdad que usa Motor.gs) — ya no hace falta tipear la cantidad a mano
    // en K3 de cada hoja individual de bono.
    const cantidadesReales = calcularCantidadesNetas();

    // 1. UNIVERSO
    const sheetUniverso = ss.getSheetByName("Universo");
    if (!sheetUniverso) throw new Error("No se encontró la hoja 'Universo'.");
    const datosUniverso = sheetUniverso.getRange("A2:U" + sheetUniverso.getLastRow()).getValues();
 const infoExtra = {};
const tickersUniverso = [];

datosUniverso.forEach(fila => {
    const tickerUSD = fila[0], tickerARS = fila[1];
    if (tickerUSD && tickerARS) {
        tickersUniverso.push([tickerUSD, tickerARS]);
        infoExtra[tickerUSD] = {
            compra: (typeof fila[2] === 'number' && fila[2] > 0) ? fila[2] : null,
            venta: (typeof fila[3] === 'number' && fila[3] > 0) ? fila[3] : null,
            empresa: fila[4] || '',
            vencimiento: fila[5] ? Math.round(cleanNum(fila[5])) : '',
            calificacion: fila[6] || '',
            tipo: fila[20] === 'Bonos' ? 'Bono' : 'ONs' // Universo!U (índice 20) — misma convención que Cartera!B
        };
    }
});

    // 2. Traducción pesos -> nombre de hoja individual (dólar)
    //    Combina Precios!S:T (lo que ya tenés en cartera) + Universo!S:T
    //    (lo que estás mirando/evaluando)
    const traduccion = {};
    [HOJAS.PRECIOS, "Universo"].forEach(nombreHoja => {
      const h = ss.getSheetByName(nombreHoja);
      if (!h) return;
      const rango = h.getRange("S2:T" + h.getLastRow()).getValues();
      rango.forEach(r => { if (r[0] && r[1]) traduccion[r[0]] = r[1]; });
    });

    // 3. Precios de mercado y macro (data912 + dolarapi + riesgo país)
    const preciosMercado = {};
    let mep = 0, ccl = 0, riesgoPais = 0, riesgoPaisVar = 0;

    const urls = [
      "https://data912.com/live/arg_bonds",
      "https://data912.com/live/arg_corp",
      "https://api.argentinadatos.com/v1/finanzas/indices/riesgo-pais"
    ];
    const requests = urls.map(url => ({ url: url, muteHttpExceptions: true }));
    let responses = [];
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      Logger.log("Error fetching APIs: " + e.toString());
      responses = requests.map(() => ({ getContentText: () => "[]" }));
    }

    try {
      const jsonBonos = JSON.parse(responses[0].getContentText());
      if (Array.isArray(jsonBonos)) jsonBonos.forEach(b => { preciosMercado[b.symbol] = b; });
    } catch (e) { Logger.log("Error Bonos: " + e.toString()); }

    try {
      const jsonONs = JSON.parse(responses[1].getContentText());
      if (Array.isArray(jsonONs)) jsonONs.forEach(o => { preciosMercado[o.symbol] = o; });
    } catch (e) { Logger.log("Error ONs: " + e.toString()); }

    try {
      const respRP = JSON.parse(responses[2].getContentText());
      if (Array.isArray(respRP) && respRP.length > 1) {
        const ultimo = respRP[respRP.length - 1], anteultimo = respRP[respRP.length - 2];
        riesgoPais = ultimo.valor;
        riesgoPaisVar = ultimo.valor - anteultimo.valor;
      } else if (Array.isArray(respRP) && respRP.length > 0) {
        riesgoPais = respRP[respRP.length - 1].valor;
      }
    } catch (e) { Logger.log("Error Riesgo País: " + e.toString()); }

    // CCL: se lee directo de Precios!B4 (fuente única de Titanium) — se
    // elimina el cálculo roto vía GGAL/GGAL.BA, que nunca funcionaba.
    try {
      const val = ss.getSheetByName(HOJAS.PRECIOS).getRange("B4").getValue();
      if (typeof val === 'number' && val > 500) ccl = val;
    } catch (e) { }

    // MEP: vía AL30/AL30D si están disponibles en la data de mercado
    if (preciosMercado["AL30"] && preciosMercado["AL30D"] && preciosMercado["AL30D"].c > 0) {
      mep = preciosMercado["AL30"].c / preciosMercado["AL30D"].c;
    } else if (ccl > 0) {
      mep = ccl / 1.025; // fallback inverso, si no hay data de AL30
    }

    const brecha = (mep > 0 && ccl > 0) ? ((ccl / mep) - 1) * 100 : 0;

    // 4. Procesamiento del universo completo
    const hoy = new Date();
    const settlementDate = getSettlementDateBombonera(hoy);
    const resultadoFinal = [];

    let totalValuadoGlobalUSD = 0, sumaTirPonderadaGlobalUSD = 0, sumaDurPonderadaGlobalUSD = 0;

    tickersUniverso.forEach(par => {
      const tickerD = par[0], tickerC = par[1];
      const instrumentoD = preciosMercado[tickerD], instrumentoC = preciosMercado[tickerC];
      if (!instrumentoD && !instrumentoC) return;

      const nombreHojaBono = traduccion[tickerC] || tickerD; // pesos -> nombre real de pestaña
      const info = infoExtra[tickerD] || {};

      let bonoFinal = {
        tickerUSD: tickerD, tickerARS: tickerC,
        tipo: info.tipo || 'ONs',
        precioUSD: instrumentoD ? instrumentoD.c : 0, var_diariaUSD: instrumentoD ? instrumentoD.pct_change : 0, volumenUSD: instrumentoD ? instrumentoD.v : 0,
        precioARS: instrumentoC ? instrumentoC.c : 0, var_diariaARS: instrumentoC ? instrumentoC.pct_change : 0, volumenARS: instrumentoC ? instrumentoC.v : 0,
        tirUSD: 'N/A', modifiedDurationUSD: 'N/A',
        mep: 0, flujosFuturos: [], nominalesUSD: 0, hasFlow: false,
        targetCompra: info.compra, targetVenta: info.venta, empresa: info.empresa,
        vencimiento: info.vencimiento, calificacion: info.calificacion
      };

      if (bonoFinal.precioUSD === 0 && bonoFinal.precioARS > 0 && mep > 0) {
        bonoFinal.precioUSD = bonoFinal.precioARS / mep;
      }
      bonoFinal.mep = (bonoFinal.precioUSD > 0 && bonoFinal.precioARS > 0) ? bonoFinal.precioARS / bonoFinal.precioUSD : mep;

      const lecturaFlujos = leerFlujosPorCada100(nombreHojaBono, settlementDate, cantidadesReales);
      bonoFinal.nominalesUSD = lecturaFlujos.tenenciaReal;
      // Copiamos también a nominalesARS: la cantidad de nominales no cambia
      // según en qué moneda estés mirando el precio — así el toggle ARS/USD
      // del frontend no se queda sin dato (ver fix en IndexBombonera.html).
      bonoFinal.nominalesARS = lecturaFlujos.tenenciaReal;
      bonoFinal.hasFlow = lecturaFlujos.flujos.length > 0;
      bonoFinal.flujosFuturos = lecturaFlujos.flujos;

      if (bonoFinal.hasFlow && bonoFinal.precioUSD > 0) {
        bonoFinal.tasaProximoCupon = lecturaFlujos.flujos[0].tasaInteres;
        bonoFinal.valorResidual = lecturaFlujos.flujos[0].residual;

        const valuesUSD = [-bonoFinal.precioUSD, ...lecturaFlujos.flujos.map(f => f.total)];
        const datesUSD = [settlementDate, ...lecturaFlujos.flujos.map(f => f.fechaPago)];

        // Reutiliza calcXIRR() de Motor.gs (con bisección) — formato {d,v}
        const cfs = valuesUSD.map((v, i) => ({ d: datesUSD[i], v: v }));
        const xirrCalc = calcXIRR(cfs);
        if (xirrCalc !== 0 || (valuesUSD.length > 1)) {
          bonoFinal.tirUSD = xirrCalc * 100;
          // TIR/Duración son propiedades del bono, no de la moneda en que lo
          // mires — copiamos también a la variante ARS para que el toggle
          // del frontend siempre tenga dato disponible.
          bonoFinal.tirARS = bonoFinal.tirUSD;
          bonoFinal.modifiedDurationUSD = calculateModifiedDurationBombonera(valuesUSD, datesUSD, xirrCalc);
          bonoFinal.modifiedDurationARS = bonoFinal.modifiedDurationUSD;
        }
      }

      if (bonoFinal.nominalesUSD > 0 && typeof bonoFinal.precioUSD === 'number' && bonoFinal.precioUSD > 0) {
        const valorActualUSD = (bonoFinal.nominalesUSD * bonoFinal.precioUSD) / 100;
        totalValuadoGlobalUSD += valorActualUSD;
        if (typeof bonoFinal.tirUSD === 'number') sumaTirPonderadaGlobalUSD += bonoFinal.tirUSD * valorActualUSD;
        if (typeof bonoFinal.modifiedDurationUSD === 'number') sumaDurPonderadaGlobalUSD += bonoFinal.modifiedDurationUSD * valorActualUSD;
      }

      resultadoFinal.push(bonoFinal);
    });

    const finalPayload = JSON.stringify({
      bonos: resultadoFinal,
      macro: {
        mep: Math.round(mep), ccl: Math.round(ccl),
        riesgoPais: Math.round(riesgoPais), riesgoPaisVar: riesgoPaisVar,
        brecha: brecha.toFixed(2)
      },
      globalMetrics: {
        totalValueUSD: totalValuadoGlobalUSD,
        weightedTirUSD: totalValuadoGlobalUSD > 0 ? sumaTirPonderadaGlobalUSD / totalValuadoGlobalUSD : 0,
        weightedDurationUSD: totalValuadoGlobalUSD > 0 ? sumaDurPonderadaGlobalUSD / totalValuadoGlobalUSD : 0
      }
    });

    cache.put(cacheKey, finalPayload, 300);
    return finalPayload;

  } catch (e) {
    Logger.log(e);
    return JSON.stringify({ error: e.toString() });
  }
}

// -----------------------------------------------------------------------
// Simulación rápida de precio custom (reemplaza calculateMetricsForCustomPrice)
// -----------------------------------------------------------------------
function calculateMetricsForCustomPriceBombonera(ticker, newPrice, currency) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const traduccion = {};
    [HOJAS.PRECIOS, "Universo"].forEach(nombreHoja => {
      const h = ss.getSheetByName(nombreHoja);
      if (!h) return;
      const rango = h.getRange("S2:T" + h.getLastRow()).getValues();
      rango.forEach(r => { if (r[0] && r[1]) traduccion[r[0]] = r[1]; });
    });

    const sheetUniverso = ss.getSheetByName("Universo");
    const datosUniverso = sheetUniverso.getRange("A2:B" + sheetUniverso.getLastRow()).getValues();
    let tickerC = null;
    for (const fila of datosUniverso) { if (fila[0] === ticker) { tickerC = fila[1]; break; } }
    const nombreHojaBono = traduccion[tickerC] || ticker;

    const hoy = new Date();
    const settlementDate = getSettlementDateBombonera(hoy);
    const cantidadesReales = calcularCantidadesNetas();
    const lectura = leerFlujosPorCada100(nombreHojaBono, settlementDate, cantidadesReales);

    if (newPrice <= 0 || lectura.flujos.length === 0) return { tir: 'N/A', duration: 'N/A' };

    const values = [-parseFloat(newPrice), ...lectura.flujos.map(f => f.total)];
    const dates = [settlementDate, ...lectura.flujos.map(f => f.fechaPago)];
    const cfs = values.map((v, i) => ({ d: dates[i], v: v }));
    const xirrCalc = calcXIRR(cfs);
    const duration = calculateModifiedDurationBombonera(values, dates, xirrCalc);

    return { tir: xirrCalc * 100, duration: duration };
  } catch (e) { return { error: e.toString() }; }
}

// -----------------------------------------------------------------------
// FIXTURE — reemplaza getPortfolioCashflow(), ahora lee de
// Proyeccion_Flujos_RF de Titanium en vez de recalcular desde cero.
// Solo incluye tickers con tenencia REAL (excluye teóricos de 100).
// -----------------------------------------------------------------------
function getPortfolioCashflowBombonera() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetProy = ss.getSheetByName(HOJAS.PROY);
    if (!sheetProy) return JSON.stringify({ error: "No se encontró Proyeccion_Flujos_RF" });

    const data = sheetProy.getRange("A2:E" + sheetProy.getLastRow()).getValues();
    const cashflowAgregado = {};

    data.forEach(row => {
      const ticker = row[0], fecha = row[1], renta = row[2], capital = row[3], total = row[4];
      if (!(fecha instanceof Date) || isNaN(fecha) || total <= 0) return;

      const mesAnio = `${fecha.getFullYear()}-${(fecha.getMonth() + 1).toString().padStart(2, '0')}`;
      if (!cashflowAgregado[mesAnio]) {
        cashflowAgregado[mesAnio] = { totalUSD: 0, rentaUSD: 0, capitalUSD: 0, totalARS: 0, rentaARS: 0, capitalARS: 0, tickers: new Set() };
      }
      // Proyeccion_Flujos_RF ya está en USD (Titanium trabaja todo en USD)
      cashflowAgregado[mesAnio].totalUSD += total;
      cashflowAgregado[mesAnio].rentaUSD += renta;
      cashflowAgregado[mesAnio].capitalUSD += capital;
      cashflowAgregado[mesAnio].tickers.add(ticker);
    });

    for (const mesAnio in cashflowAgregado) cashflowAgregado[mesAnio].tickers = Array.from(cashflowAgregado[mesAnio].tickers);
    return JSON.stringify(cashflowAgregado);
  } catch (e) { return JSON.stringify({ error: e.toString() }); }
}