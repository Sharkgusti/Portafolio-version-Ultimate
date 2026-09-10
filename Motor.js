// =================================================================================
// ===  TITANIUM v2 — MOTOR.GS                                                  ===
// ===  Única fuente: Log_Transacciones_TITANIUM                                ===
// =================================================================================
 
const HOJAS = {
    LOG: "Log_Transacciones_TITANIUM",
    PRECIOS: "Precios",
    CARTERA: "Cartera",
    EVO: "Evolucion cartera",
    PROY: "Proyeccion_Flujos_RF",
    HIST: "Historico_Precios",
    QUANT: "Analisis_Quant"
};
 
const HOY_SIMULADA = new Date();
HOY_SIMULADA.setHours(0, 0, 0, 0);
 
// =================================================================================
// CONSTANTES A REVISAR EN CADA INSTALACIÓN NUEVA DEL SISTEMA (ver debate
// "Portafolio Ultimate"). Los tres valores de acá abajo son específicos del
// historial de carga de datos de ESTE usuario particular — si este archivo se
// usa como base para otra cartera (por ejemplo, una copia del Sheet para otra
// persona), hay que revisar y ajustar estos tres valores a la realidad de esa
// persona antes de confiar en ningún cálculo de rendimiento o de caja.
//
//   FECHA_INICIO_CONFIABLE_DATOS: fecha a partir de la cual el registro de
//     movimientos de este usuario se considera completo y confiable. Actúa
//     como piso para: (a) el saldo de caja virtual (calcularCajaVirtual, que
//     ignora todo movimiento de caja anterior a esta fecha) y (b) los
//     períodos de rendimiento del Dashboard (mes/trim/sem/año no pueden
//     empezar antes de esta fecha). Si una instalación nueva tiene registros
//     completos desde el día 1 (por ejemplo, alguien que arranca su Log de
//     cero), esta fecha debería ser la fecha del primer movimiento real, no
//     quedarse en la de este usuario.
//
//   SALDO_INICIAL_CAJA: el saldo de caja en USD que había, reconciliado a
//     mano, justo en FECHA_INICIO_CONFIABLE_DATOS (representa la plata que
//     no quedó registrada movimiento por movimiento antes de esa fecha). En
//     una instalación nueva sin ese arrastre, debería ser 0.
// =================================================================================
const FECHA_INICIO_CONFIABLE_DATOS = new Date(2026, 1, 1); // 1/2/2026
const SALDO_INICIAL_CAJA = 14;
 
const FECHA_CORTE_CAJA = FECHA_INICIO_CONFIABLE_DATOS; // alias: mismo concepto, usado en calcularCajaVirtual()
const TENENCIA_TEORICA_BOMBONERA = 100;
 
// 1. API WEB APP
// =================================================================================
// Clave secreta para el endpoint de datos crudos (?formato=json). La Web App
// es pública y anónima (ANYONE_ANONYMOUS) — sin esta clave, cualquiera con el
// link del Dashboard podría bajarse la cartera completa con un curl, sin
// pasar por la interfaz. Cambiala por la tuya, y no la compartas.
// =================================================================================
const CLAVE_SECRETA_API = 'kO544sk6YxiI5bnPUS2_OpcU5wCu9p5S';
 
function doGet(e) {
  // Endpoint de datos crudos para consumo externo (ej. Python/Colab) — ver
  // debate "Portafolio Ultimate". No devuelve ninguna página, solo el JSON
  // que ya arma generarDatosMaestros(), tal cual, para que Python calcule
  // métricas sin reimplementar el motor de posiciones en un tercer lenguaje.
  if (e && e.parameter && e.parameter.formato === 'json') {
    if (e.parameter.clave !== CLAVE_SECRETA_API) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Clave inválida o faltante' }))
          .setMimeType(ContentService.MimeType.JSON);
    }
    const datosJson = generarDatosMaestros(); // ya es un string JSON, no hace falta re-serializar
    return ContentService.createTextOutput(datosJson)
        .setMimeType(ContentService.MimeType.JSON);
  }
 
  let template;
  let isBombonera = (e && e.parameter && e.parameter.page === 'bombonera');
 
  if (isBombonera) {
    template = HtmlService.createTemplateFromFile('IndexBombonera');
  } else {
    template = HtmlService.createTemplateFromFile('Index');
  }
 
  try {
    template.webAppUrl = ScriptApp.getService().getUrl();
  } catch (err) {
    template.webAppUrl = '';
  }
 
  return template.evaluate()
      .setTitle(isBombonera ? 'La Bombonera Digital del Contadore' : 'TITANIUM v2')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. MOTOR PRINCIPAL
function generarDatosMaestros() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const logData = ss.getSheetByName(HOJAS.LOG).getDataRange().getValues();
    const preciosData = ss.getSheetByName(HOJAS.PRECIOS).getDataRange().getValues();

    let ccl = 1000;
    try {
        const val = ss.getSheetByName(HOJAS.PRECIOS).getRange("B4").getValue();
        if (typeof val === 'number' && val > 500) ccl = val;
    } catch (e) { }

    const precios = {};
    for (let i = 1; i < preciosData.length; i++) {
        let tk = String(preciosData[i][0]).toUpperCase().trim();
        let pr = cleanNum(preciosData[i][1]);
        let moneda = String(preciosData[i][2] || "").toUpperCase().trim();
        let vr = cleanNum(preciosData[i][4]);
        if (tk && pr > 0) precios[tk] = { precio: pr, var: vr, moneda: moneda };
    }

    const logSinHeader = logData.slice(1);
    logSinHeader.sort((a, b) => new Date(a[1]) - new Date(b[1]));

    let cajaVirtual = calcularCajaVirtual(logSinHeader);

    // =============================================================================
    // PASO 3.4 de la consolidación (ver debate "Portafolio Ultimate"): el motor
    // principal del Dashboard ya NO reimplementa el loop de posiciones. La
    // posición de HOY sale de snapshotAFecha() y todo el detalle histórico
    // (cashflows, ganancias realizadas, dividendos, renta, Modified Dietz) se
    // deriva de ledgerCompleto() — ambas en MotorPosiciones.gs. Ninguna regla
    // de negocio cambia acá: Método B, tope de venta y venta-sin-stock siguen
    // siendo exactamente las mismas que ya vienen aplicando desde el Paso 3.1.
    // =============================================================================
    const snapshotHoy = snapshotAFecha(logSinHeader, HOY_SIMULADA);
    const ledger = ledgerCompleto(logSinHeader);
    let portfolio = snapshotHoy.posiciones;

    const esRVporTipo = (tipoStr) => String(tipoStr).toLowerCase().includes('cedear') || String(tipoStr).toLowerCase().includes('accion');

    // --- Clasificación auxiliar: tickers con renta/dividendo explícito en su historia ---
    let tickersConRentaExplicita = new Set();
    ledger.transacciones.forEach(t => {
        if (t.movimiento.includes('dividendo') || t.movimiento.includes('renta') || t.movimiento.includes('interes')) {
            tickersConRentaExplicita.add(t.ticker);
        }
    });

    // --- Derivación 1: cashflows agrupados (RV / RF / Total) y por ticker, + ganancia realizada RV ---
    let cfRV = [], cfRF = [], cfTOTAL = [];
    let cashflowsPorTicker = {};
    let realizedPL_RV = 0;

    ledger.transacciones.forEach(t => {
        const esRV = esRVporTipo(t.tipo);
        if (!cashflowsPorTicker[t.ticker]) cashflowsPorTicker[t.ticker] = [];

        if (t.cashflow) {
            cfTOTAL.push({ d: t.fecha, v: t.cashflow });
            if (esRV) cfRV.push({ d: t.fecha, v: t.cashflow });
            else cfRF.push({ d: t.fecha, v: t.cashflow });
            cashflowsPorTicker[t.ticker].push({ d: t.fecha, v: t.cashflow });
        }

        const esVentaLike = t.movimiento.includes('venta') || t.movimiento.includes('rescate') || t.movimiento.includes('canje_salida');
        const esRentaLike = t.movimiento.includes('dividendo') || t.movimiento.includes('renta') || t.movimiento.includes('interes');
        if (esRV && (esVentaLike || esRentaLike)) {
            realizedPL_RV += t.gananciaRealizada;
        }
    });

    // --- Derivación 2: estadísticas de dividendos/renta por ticker y por año, y
    //     ganancia realizada por ticker/año (misma lógica de negocio de siempre,
    //     incluida la regla de tickersConRentaExplicita para la ganancia extra
    //     de amortización, y el FIX que suma esa ganancia extra a
    //     gananciaRealizadaPorTickerAnio SIEMPRE, sea o no cupón explícito). ---
    let statsDivs = { tickers: {}, years: {}, total: 0 };
    let statsRenta = { tickers: {}, years: {}, total: 0 };
    let gananciaRealizadaPorTickerAnio = {};
    let fechaHaceUnAnio = new Date(); fechaHaceUnAnio.setFullYear(fechaHaceUnAnio.getFullYear() - 1);
    let rentaHistorica12m = {};

    ledger.transacciones.forEach(t => {
        const year = t.anio;
        const esRF = !esRVporTipo(t.tipo);

        if (t.movimiento.includes('dividendo')) {
            statsDivs.tickers[t.ticker] = (statsDivs.tickers[t.ticker] || 0) + t.montoUSD;
            statsDivs.years[year] = (statsDivs.years[year] || 0) + t.montoUSD;
            statsDivs.total += t.montoUSD;
            if (t.fecha >= fechaHaceUnAnio) rentaHistorica12m[t.ticker] = (rentaHistorica12m[t.ticker] || 0) + t.montoUSD;

        } else if (t.movimiento.includes('renta') || t.movimiento.includes('interes')) {
            statsRenta.tickers[t.ticker] = (statsRenta.tickers[t.ticker] || 0) + t.montoUSD;
            statsRenta.years[year] = (statsRenta.years[year] || 0) + t.montoUSD;
            statsRenta.total += t.montoUSD;
            if (t.fecha >= fechaHaceUnAnio) rentaHistorica12m[t.ticker] = (rentaHistorica12m[t.ticker] || 0) + t.montoUSD;

            if (!gananciaRealizadaPorTickerAnio[t.ticker]) gananciaRealizadaPorTickerAnio[t.ticker] = {};
            gananciaRealizadaPorTickerAnio[t.ticker][year] = (gananciaRealizadaPorTickerAnio[t.ticker][year] || 0) + t.montoUSD;

        } else if (t.movimiento.includes('amortiza')) {
            if (t.gananciaRealizada > 0) {
                if (tickersConRentaExplicita.has(t.ticker)) {
                    statsRenta.tickers[t.ticker] = (statsRenta.tickers[t.ticker] || 0) + t.gananciaRealizada;
                    statsRenta.years[year] = (statsRenta.years[year] || 0) + t.gananciaRealizada;
                    statsRenta.total += t.gananciaRealizada;
                }
                // FIX (preservado del sistema anterior): esto va SIEMPRE, sea o no
                // un ticker con cupón explícito — es plata ya cobrada en efectivo
                // (Realizada), sin importar a qué bucket del total termine yendo.
                if (!gananciaRealizadaPorTickerAnio[t.ticker]) gananciaRealizadaPorTickerAnio[t.ticker] = {};
                gananciaRealizadaPorTickerAnio[t.ticker][year] = (gananciaRealizadaPorTickerAnio[t.ticker][year] || 0) + t.gananciaRealizada;
            }

        } else if (t.movimiento.includes('venta') || t.movimiento.includes('rescate') || t.movimiento.includes('canje_salida')) {
            if (esRF && !tickersConRentaExplicita.has(t.ticker)) {
                if (!gananciaRealizadaPorTickerAnio[t.ticker]) gananciaRealizadaPorTickerAnio[t.ticker] = {};
                gananciaRealizadaPorTickerAnio[t.ticker][year] = (gananciaRealizadaPorTickerAnio[t.ticker][year] || 0) + t.gananciaRealizada;
            }
        }
    });

    // --- Derivación 3: costo vivo RF a través del tiempo, para el yield ponderado ---
    let totalCostoVivoRF = 0;
    let costoEventsRF = [];
    let ultimoCostoConocido = {};

    ledger.transacciones.forEach(t => {
        const esRF = !esRVporTipo(t.tipo);
        const previo = (ultimoCostoConocido[t.ticker] !== undefined) ? ultimoCostoConocido[t.ticker] : 0;
        const delta = t.costoDespues - previo;
        ultimoCostoConocido[t.ticker] = t.costoDespues;

        const esEventoDeCosto = t.movimiento.includes('compra') || t.movimiento.includes('aporte') ||
            t.movimiento.includes('suscripcion') || t.movimiento.includes('canje_entrada') ||
            t.movimiento.includes('venta') || t.movimiento.includes('rescate') || t.movimiento.includes('canje_salida') ||
            t.movimiento.includes('amortiza');

        if (esRF && esEventoDeCosto) {
            totalCostoVivoRF += delta;
            costoEventsRF.push({ d: t.fecha, c: totalCostoVivoRF });
        }
    });
    costoEventsRF.sort((a, b) => a.d - b.d);

    const hPreciosCompleto = leerHistPreciosCompleto(ss.getSheetByName(HOJAS.HIST));
    const hPrecios = submuestrearHistPrecios(hPreciosCompleto, 150);

    // --- Derivación 4: Modified Dietz para bonos/ONs sin cupón explícito ---
    let unitEventsPorTicker = {};
    ledger.transacciones.forEach(t => {
        const esEventoDeUnidades = t.movimiento.includes('compra') || t.movimiento.includes('aporte') ||
            t.movimiento.includes('suscripcion') || t.movimiento.includes('canje_entrada') ||
            t.movimiento.includes('venta') || t.movimiento.includes('rescate') || t.movimiento.includes('canje_salida') ||
            t.movimiento.includes('split');
        if (!esEventoDeUnidades) return;
        if (!unitEventsPorTicker[t.ticker]) unitEventsPorTicker[t.ticker] = [];
        unitEventsPorTicker[t.ticker].push({ d: t.fecha, q: t.qDespues });
    });

    Object.keys(unitEventsPorTicker).forEach(tk => {
        if (tickersConRentaExplicita.has(tk)) return;
        const p = portfolio[tk];
        if (!p) return;
        const tipoNorm = normalizar(p.tipo);
        const esRFTk = tipoNorm !== 'Cedear' && tipoNorm !== 'Acciones';
        if (!esRFTk) return;

        const eventosQ = unitEventsPorTicker[tk].slice().sort((a, b) => a.d - b.d);
        if (eventosQ.length === 0) return;

        const preciosHist = (hPreciosCompleto[tk] || []).map(pt => ({ d: new Date(pt[0]), p: pt[1] })).sort((a, b) => a.d - b.d);

        const buscarPrecio = (fecha) => {
            let ultimo = null;
            for (let i = 0; i < preciosHist.length; i++) {
                if (preciosHist[i].d <= fecha) ultimo = preciosHist[i].p;
                else break;
            }
            return ultimo;
        };
        const buscarUnidades = (fecha) => {
            let ultimo = 0;
            for (let i = 0; i < eventosQ.length; i++) {
                if (eventosQ[i].d <= fecha) ultimo = eventosQ[i].q;
                else break;
            }
            return ultimo;
        };

        const primerAnio = eventosQ[0].d.getFullYear();
        const ultimoAnio = HOY_SIMULADA.getFullYear();

        for (let anio = primerAnio; anio <= ultimoAnio; anio++) {
            const inicioAnio = new Date(anio, 0, 1);
            const finAnio = new Date(anio, 11, 31);
            const corte = (anio === ultimoAnio) ? HOY_SIMULADA : finAnio;

            const unidadesInicio = buscarUnidades(new Date(inicioAnio.getTime() - 86400000));
            const unidadesFin = buscarUnidades(corte);
            const precioInicio = (unidadesInicio > 0) ? buscarPrecio(inicioAnio) : 0;
            const precioFin = (unidadesFin > 0) ? buscarPrecio(corte) : 0;

            if ((unidadesInicio > 0 && precioInicio === null) || (unidadesFin > 0 && precioFin === null)) continue;

            const valorInicio = unidadesInicio * (precioInicio || 0);
            const valorFin = unidadesFin * (precioFin || 0);

            let comprasAnio = 0, rescatesAnio = 0;
            ledger.transacciones.forEach(t => {
                if (t.ticker !== tk || t.anio !== anio) return;
                if (t.movimiento.includes('compra') || t.movimiento.includes('aporte') || t.movimiento.includes('suscripcion') || t.movimiento.includes('canje_entrada')) comprasAnio += t.montoUSD;
                else if (t.movimiento.includes('venta') || t.movimiento.includes('rescate') || t.movimiento.includes('amortiza') || t.movimiento.includes('canje_salida')) rescatesAnio += t.montoUSD;
            });

            if (unidadesInicio === 0 && unidadesFin === 0 && comprasAnio === 0 && rescatesAnio === 0) continue;

            const ganancia = valorFin - valorInicio - comprasAnio + rescatesAnio;
            if (Math.abs(ganancia) < 0.01) continue;

            statsRenta.tickers[tk] = (statsRenta.tickers[tk] || 0) + ganancia;
            statsRenta.years[anio] = (statsRenta.years[anio] || 0) + ganancia;
            statsRenta.total += ganancia;
        }
    });

    // =============================================================================
    // A partir de acá, la función sigue exactamente igual que antes del Paso 3.4:
    // formateo de resultados, valuación a precio de hoy, armado del JSON final.
    // =============================================================================
    const formatObj = (obj, total) => ({
        total: total,
        porTicker: Object.keys(obj.tickers).map(k => ({ ticker: k, monto: obj.tickers[k] })).sort((a, b) => b.monto - a.monto),
        porAno: Object.keys(obj.years).map(k => ({ ano: parseInt(k), monto: obj.years[k] })).sort((a, b) => b.ano - a.ano)
    });
    const resDivs = formatObj(statsDivs, statsDivs.total);

    let realizadoPorAnio = {};
    Object.keys(gananciaRealizadaPorTickerAnio).forEach(tk => {
        Object.keys(gananciaRealizadaPorTickerAnio[tk]).forEach(anioStr => {
            let anio = parseInt(anioStr);
            realizadoPorAnio[anio] = (realizadoPorAnio[anio] || 0) + gananciaRealizadaPorTickerAnio[tk][anioStr];
        });
    });

    const resRenta = {
        total: statsRenta.total,
        porTicker: Object.keys(statsRenta.tickers).map(k => ({ ticker: k, monto: statsRenta.tickers[k] })).sort((a, b) => b.monto - a.monto),
        porAno: Object.keys(statsRenta.years).map(k => {
            let anio = parseInt(k);
            let monto = statsRenta.years[anio];
            let cp = calcCostoPromedioPonderado(costoEventsRF, anio, HOY_SIMULADA);
            let yieldCrudo = cp.costoProm > 0 ? monto / cp.costoProm : 0;
            let yieldFinal = (cp.esParcial && cp.dias > 0) ? yieldCrudo * 365 / cp.dias : yieldCrudo;
            let realizada = realizadoPorAnio[anio] || 0;
            let noRealizada = monto - realizada;
            return { ano: anio, monto: monto, costoPromedio: cp.costoProm, yield: yieldFinal, parcial: cp.esParcial, realizada: realizada, noRealizada: noRealizada };
        }).sort((a, b) => b.ano - a.ano)
    };

    let rentaFutura12m = {};
    let fechaDentroUnAnio = new Date(); fechaDentroUnAnio.setFullYear(fechaDentroUnAnio.getFullYear() + 1);
    try {
        const proyData = ss.getSheetByName(HOJAS.PROY).getDataRange().getValues();
        for (let r = 1; r < proyData.length; r++) {
            let tk = String(proyData[r][0]).trim().toUpperCase();
            let fecha = new Date(proyData[r][1]);
            let montoRenta = cleanNum(proyData[r][2]);
            if (fecha <= fechaDentroUnAnio && montoRenta > 0) {
                rentaFutura12m[tk] = (rentaFutura12m[tk] || 0) + montoRenta;
            }
        }
    } catch (e) { }

    let lista = [], valTotalUSD = 0, costoTotal = 0, dist = {};
    let valFinalRV = 0, valFinalRF = 0;
    let valorTotalRV = 0, costoTotalRV = 0;
    let rentaAnualTotalEstimada = 0;

    Object.keys(portfolio).forEach(tk => {
        let p = portfolio[tk];
        if (p.q <= 0.01) return;

        let pxData = precios[tk] || { precio: 0, var: 0, moneda: "" };
        let precioARS = pxData.precio;
        let monedaFlag = pxData.moneda;
        let valUSD = 0;
        let tipoNorm = normalizar(p.tipo);

        if (tipoNorm === 'Bonos' || tipoNorm === 'ONs') {
            if (monedaFlag === 'USD') valUSD = precioARS * p.q;
            else if (monedaFlag === 'PESOS') valUSD = (precioARS / ccl) * p.q;
            else valUSD = (precioARS / ccl / 100) * p.q;
        } else if (tipoNorm === 'Cedear' || tipoNorm === 'Acciones') {
            valUSD = (precioARS / ccl) * p.q;
        } else {
            if (monedaFlag === 'USD') valUSD = precioARS * p.q;
            else valUSD = (precioARS / ccl) * p.q;
        }

        if (!cashflowsPorTicker[tk]) cashflowsPorTicker[tk] = [];
        cashflowsPorTicker[tk].push({ d: HOY_SIMULADA, v: valUSD });
        valTotalUSD += valUSD;
        costoTotal += p.costo;

        let esRV = tipoNorm === 'Cedear' || tipoNorm === 'Acciones';
        if (esRV) { valorTotalRV += valUSD; costoTotalRV += p.costo; valFinalRV += valUSD; }
        else { valFinalRF += valUSD; }

        dist[tipoNorm] = (dist[tipoNorm] || 0) + valUSD;

        let rentaEstimada = (tipoNorm === 'Bonos' || tipoNorm === 'ONs') ? (rentaFutura12m[tk] || 0) : (rentaHistorica12m[tk] || 0);
        rentaAnualTotalEstimada += rentaEstimada;

        let roi = (p.costoOriginal > 0) ? (valUSD + p.cobrado) / p.costoOriginal - 1 : 0;
        let xirrCalculado = calcXIRR(cashflowsPorTicker[tk]);
        let diasTenencia = p.wDate ? Math.floor((HOY_SIMULADA.getTime() - p.wDate) / 86400000) : 0;

        lista.push({
            ticker: tk, tipo: tipoNorm,
            unidades: p.q, costoUSD: p.costo, valorUSD: valUSD,
            ppc: (p.q > 0) ? (p.costo / p.q) : 0,
            variacionDiaria: pxData.var,
            plPorcentaje: roi,
            xirr: xirrCalculado,
            yoc: (p.costo > 0) ? rentaEstimada / p.costo : 0,
            currentYield: (valUSD > 0) ? rentaEstimada / valUSD : 0,
            totalCobrado: p.cobrado,
            diasTenencia: diasTenencia
        });
    });

    if (cajaVirtual !== 0) {
        lista.push({
            ticker: 'LIQUIDEZ_AUTO', tipo: 'Liquidez',
            unidades: cajaVirtual, costoUSD: cajaVirtual, valorUSD: cajaVirtual,
            variacionDiaria: 0, plPorcentaje: 0, xirr: 0, yoc: 0, currentYield: 0, totalCobrado: 0, diasTenencia: ''
        });
        valTotalUSD += cajaVirtual;
        dist['Liquidez'] = (dist['Liquidez'] || 0) + cajaVirtual;
    }

    lista.forEach(x => x.porcentajeCartera = (valTotalUSD > 0) ? x.valorUSD / valTotalUSD : 0);

    if (valFinalRV > 0) cfRV.push({ d: HOY_SIMULADA, v: valFinalRV });
    if (valFinalRF > 0) cfRF.push({ d: HOY_SIMULADA, v: valFinalRF });
    if (valTotalUSD > 0) cfTOTAL.push({ d: HOY_SIMULADA, v: valTotalUSD });

    const hojaEvo = ss.getSheetByName(HOJAS.EVO);
    let valorAyerUSD = obtenerValorAyer(hojaEvo);
    let variacionDiariaReal = (valorAyerUSD > 0) ? (valTotalUSD - valorAyerUSD) : 0;
    let variacionPorcentual = (valorAyerUSD > 0) ? (variacionDiariaReal / valorAyerUSD) : 0;

    const qMetrics = leerQuant(ss.getSheetByName(HOJAS.QUANT));
    const histEvo = leerEvo(hojaEvo);
    const flujosFut = leerProyecciones(ss.getSheetByName(HOJAS.PROY));
    const chartAcum = procesarFlujosAcumulativosMensuales(ss.getSheetByName(HOJAS.PROY));
    const chartSem = procesarFlujosSemestrales(ss.getSheetByName(HOJAS.PROY));

    const spyData = precios["INDICE_SPY"] || { var: 0 };
    const alpha = variacionPorcentual - (spyData.var / 100);
    const rendimientoYTD = calcularRendimientoLimpioYTD(logSinHeader, portfolio, precios, ccl, HOY_SIMULADA, hPreciosCompleto, histEvo);
    const inicioDatosRendimiento = new Date(2026, 1, 1);
    const inicioPeriodo = (dias) => {
        const fecha = new Date(HOY_SIMULADA);
        fecha.setDate(fecha.getDate() - dias);
        return fecha < inicioDatosRendimiento ? new Date(inicioDatosRendimiento) : fecha;
    };
    const rendimientoPeriodos = {
        mes: calcularRendimientoLimpioYTD(logSinHeader, portfolio, precios, ccl, HOY_SIMULADA, hPreciosCompleto, histEvo, inicioPeriodo(30)),
        trim: calcularRendimientoLimpioYTD(logSinHeader, portfolio, precios, ccl, HOY_SIMULADA, hPreciosCompleto, histEvo, inicioPeriodo(90)),
        sem: calcularRendimientoLimpioYTD(logSinHeader, portfolio, precios, ccl, HOY_SIMULADA, hPreciosCompleto, histEvo, inicioPeriodo(180)),
        anio: calcularRendimientoLimpioYTD(logSinHeader, portfolio, precios, ccl, HOY_SIMULADA, hPreciosCompleto, histEvo, inicioPeriodo(365)),
        max: rendimientoYTD
    };

    const logParaFrontend = logSinHeader.map(row => ({
        fecha: row[1] instanceof Date ? row[1].toISOString() : String(row[1]),
        ticker: String(row[3] || '').toUpperCase().trim(),
        tipo: String(row[4] || ''),
        mov: String(row[5] || '').toLowerCase().trim(),
        montoUSD: Math.abs(cleanNum(row[10]))
    }));

    return JSON.stringify({
        kpis: {
            valorTotal: { ars: valTotalUSD * ccl, usd: valTotalUSD },
            valorRV: valorTotalRV,
            valorRF: valFinalRF,
            variacionDiariaUSD: variacionDiariaReal,
            variacionDiariaPorc: variacionPorcentual,
            gananciasRealizadas: { usd: realizedPL_RV },
            gananciasNoRealizadas: { usd: valorTotalRV - costoTotalRV },
            tirVariable: calcXIRR(cfRV),
            tirFija: calcXIRR(cfRF),
            tirTotal: calcXIRR(cfTOTAL),
            alphaSpy: alpha,
            rentaAnualEstimada: rentaAnualTotalEstimada,
            liquidezDisponible: cajaVirtual,
            currentYieldPortfolio: (valTotalUSD > 0) ? (rentaAnualTotalEstimada / valTotalUSD) : 0,
            cagrHistorico: ((histEvo.fechas.length > 30 && histEvo.valores[0] > 0) ? (Math.pow(histEvo.valores[histEvo.valores.length - 1] / histEvo.valores[0], 1 / ((new Date(histEvo.fechas[histEvo.fechas.length - 1]) - new Date(histEvo.fechas[0])) / (1000 * 60 * 60 * 24 * 365.25))) - 1) : 0),
            yocPromedioPonderado: (costoTotal > 0) ? (rentaAnualTotalEstimada / costoTotal) : 0,
            sortinoRatio: calcSortino(histEvo.valores),
            dividendGrowthRate: calcDGR(statsDivs),
            paybackYears: ((statsDivs.total + statsRenta.total) > 0) ? (costoTotal / (statsDivs.total + statsRenta.total)) : 99,
            nextPay: calcNextPay(flujosFut.porFecha),
            rendimientoLimpio2026: rendimientoYTD,
            rendimientoPeriodos: rendimientoPeriodos
        },
        metricasQuant: qMetrics,
        distribucionTipos: dist,
        carteraDetallada: lista,
        mejoresPosiciones: { Cedear: top(lista, "Cedear"), ONs: top(lista, "ONs"), Bonos: top(lista, "Bonos"), Acciones: top(lista, "Acciones") },
        evolucionHistorica: histEvo,
        resumenDividendos: resDivs,
        resumenRentaFija: resRenta,
        flujoFuturoRentaFija: flujosFut,
        flujoAcumulativoMensual: chartAcum,
        flujoSemestral: chartSem,
        historicoPreciosActivos: hPrecios,
        logTransacciones: logParaFrontend
    });
}


// =================================================================================
// ===  MOTOR.GS — PARTE 2: FUNCIONES AUXILIARES                                ===
// =================================================================================

function calcularRendimientoLimpioYTD(logSinHeader, portfolio, precios, ccl, hoy, hPreciosCompleto, histEvo, fechaInicioOpcional) {
    const anioActual = hoy.getFullYear();
    const FECHA_CORTE = fechaInicioOpcional ? new Date(fechaInicioOpcional) : ((anioActual === 2026) ? new Date(2026, 1, 1) : new Date(anioActual, 0, 1));
    const fechaCorteMs = FECHA_CORTE.getTime();
    const fechaCorteStr = Utilities.formatDate(FECHA_CORTE, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    if (!logSinHeader || logSinHeader.length === 0) {
        return {
            anio: anioActual,
            fechaCorteStr: fechaCorteStr,
            tirTotal: 0, tirVariable: 0, tirFija: 0, aporteNetoUSD: 0,
            retAbsolutoTotal: 0, retAbsolutoVariable: 0, retAbsolutoFija: 0, gananciaNetaTotalUSD: 0
        };
    }
    let aporteNetoUSD = 0;
    let comprasNetasRV = 0;
    let comprasNetasRF = 0;

    // =============================================================================
    // PASO 3.3 de la consolidación (ver debate "Portafolio Ultimate"): la
    // tenencia a la fecha de corte (tenenciasCorte) ya NO se calcula con un
    // mini-loop propio — sale de snapshotAFecha() (MotorPosiciones.gs), que
    // aplica las mismas reglas que el resto del motor (Método B, tope de
    // venta, venta sin stock previo). El tracking de ultimoPrecioUSD se
    // mantiene igual: es lógica de PRECIOS (fallback cuando no hay histórico
    // al corte), no de posiciones, y no pertenece al motor de posiciones.
    // =============================================================================
    let ultimoPrecioUSD = {};

    logSinHeader.forEach(row => {
        let fechaObj = row[1];
        if (!(fechaObj instanceof Date)) fechaObj = new Date(row[1]);
        if (isNaN(fechaObj.getTime())) return;

        const ticker = String(row[3]).toUpperCase().trim();
        const pxOrig = cleanNum(row[7]);
        const tc = cleanNum(row[9]) || 1;
        const monOrig = String(row[8] || "").toUpperCase().trim();

        if (!ticker || ticker === 'USD' || ticker === 'CASH') return;

        let pxUSD = 0;
        if (monOrig === 'USD') pxUSD = pxOrig;
        else if (monOrig === 'ARS' || monOrig === 'PESOS') pxUSD = pxOrig / tc;
        if (pxUSD > 0) ultimoPrecioUSD[ticker] = pxUSD;
    });

    const snapshotCorte = snapshotAFecha(logSinHeader, FECHA_CORTE);
    let tenenciasCorte = {};
    Object.keys(snapshotCorte.posiciones).forEach(tk => {
        tenenciasCorte[tk] = { q: snapshotCorte.posiciones[tk].q, tipo: snapshotCorte.posiciones[tk].tipo };
    });
    
    const buscarPrecioAlCorte = (tk) => {
        if (!hPreciosCompleto || !hPreciosCompleto[tk]) return null;
        const list = hPreciosCompleto[tk];
        let px = null;
        for (let i = 0; i < list.length; i++) {
            let t = list[i][0];
            let tMs = (typeof t === 'number') ? t : new Date(t).getTime();
            if (tMs <= fechaCorteMs) px = list[i][1];
            else break;
        }
        return px;
    };
    let valInicialRV = 0, valInicialRFDirecto = 0;
    Object.keys(tenenciasCorte).forEach(tk => {
        let q = tenenciasCorte[tk].q;
        if (q <= 0.01) return;
        let tNorm = normalizar(tenenciasCorte[tk].tipo);
        let esRV = (tNorm === 'Cedear' || tNorm === 'Acciones');

        let pxHist = buscarPrecioAlCorte(tk);
        let valUSD;

        if (pxHist !== null && pxHist > 0) {
            // FIX: Historico_Precios ya guarda el precio normalizado en USD
            // por unidad (misma convención que usa toda la Renta Implícita
            // / Modified Dietz) — antes se lo dividía por el CCL una segunda
            // vez acá, achicando el valor inicial de RV por un factor de
            // ~1500x (el CCL) y corrompiendo en cascada tanto el P/L% de RV
            // (denominador casi cero -> porcentaje absurdo) como, por el
            // cálculo de RF como residuo (Total - RV - Caja), el XIRR/P&L
            // de Renta Fija (absorbía de más el valor inicial que le
            // correspondía a RV).
            valUSD = pxHist * q;
        } else {
            // Sin histórico disponible para este ticker al corte: fallback
            // al precio ARS actual (o al costo promedio como último
            // recurso), convertido normal por CCL — esta rama sí estaba bien.
            let pxData = precios[tk] || { precio: 0, moneda: "" };
            let pxARS = pxData.precio;
            if (!pxARS || pxARS <= 0) pxARS = (ultimoPrecioUSD[tk] || 0) * ccl;
            if (!pxARS || pxARS <= 0) {
                if (portfolio[tk] && portfolio[tk].q > 0) pxARS = (portfolio[tk].costo / portfolio[tk].q) * ccl;
            }
            valUSD = (pxARS / ccl) * q;
        }

        if (esRV) valInicialRV += valUSD;
        else valInicialRFDirecto += valUSD;
    });


    // Final Portfolio Valuation at TODAY
    let valFinalRV = 0, valFinalRF = 0, valFinalTotal = 0;
    Object.keys(portfolio).forEach(tk => {
        let p = portfolio[tk];
        if (p.q <= 0.01) return;
        let pxData = precios[tk] || { precio: 0, moneda: "" };
        let pxARS = pxData.precio;
        let mon = pxData.moneda;
        let tNorm = normalizar(p.tipo);
        let valUSD = 0;

        if (tNorm === 'Bonos' || tNorm === 'ONs') {
            if (mon === 'USD') valUSD = pxARS * p.q;
            else if (mon === 'PESOS') valUSD = (pxARS / ccl) * p.q;
            else valUSD = (pxARS / ccl / 100) * p.q;
        } else if (tNorm === 'Cedear' || tNorm === 'Acciones') {
            valUSD = (pxARS / ccl) * p.q;
        } else {
            if (mon === 'USD') valUSD = pxARS * p.q;
            else valUSD = (pxARS / ccl) * p.q;
        }

        let esRV = (tNorm === 'Cedear' || tNorm === 'Acciones');
        if (esRV) valFinalRV += valUSD;
        else valFinalRF += valUSD;
        valFinalTotal += valUSD;
    });

    let cajaVirtualActual = calcularCajaVirtual(logSinHeader);
    valFinalTotal += cajaVirtualActual;

    // Lookup exact baseline total portfolio value from histEvo on cutoff date
    let valInicialTotal = 0;
    if (histEvo && histEvo.fechas) {
        for (let i = 0; i < histEvo.fechas.length; i++) {
            let dMs = new Date(histEvo.fechas[i]).getTime();
            if (dMs <= fechaCorteMs) valInicialTotal = histEvo.valores[i];
            else break;
        }
    }

    const cajaInicialCorte = calcularCajaVirtualHasta(logSinHeader, FECHA_CORTE);
    if (valInicialTotal <= 0) valInicialTotal = Math.max(1000, valFinalTotal * 0.88);
    if (valInicialRV <= 0) valInicialRV = Math.max(100, valFinalRV * 0.828);

    // Evolución cartera es la fuente de verdad del total. RF se obtiene como
    // residual contable para que Total = RV + RF + Caja en cada fecha de corte.
    // Los precios históricos de RF quedan como respaldo ante un residual inválido.
    let valInicialRF = valInicialTotal - valInicialRV - cajaInicialCorte;
    if (valInicialRF <= 0) valInicialRF = valInicialRFDirecto;
    if (valInicialRF <= 0) valInicialRF = Math.max(100, valFinalRF * 0.92);

    let bridgeAportes = 0, bridgeRetiros = 0, bridgeDividendos = 0, bridgeRenta = 0, bridgeAmortizaciones = 0;
    let cfTOTAL = [{ d: FECHA_CORTE, v: -valInicialTotal }];
    let cfRV = [{ d: FECHA_CORTE, v: -valInicialRV }];
    let cfRF = [{ d: FECHA_CORTE, v: -valInicialRF }];

    logSinHeader.forEach(row => {
        let fechaObj = row[1];
        if (!(fechaObj instanceof Date)) fechaObj = new Date(row[1]);
        if (isNaN(fechaObj.getTime()) || fechaObj.getTime() <= fechaCorteMs) return;

        const ticker = String(row[3]).toUpperCase().trim();
        const tipoStr = String(row[4]);
        const mov = String(row[5]).toLowerCase().trim();
        const montoUSD = Math.abs(cleanNum(row[10]));
        const esCash = (ticker === 'USD' || ticker === 'CASH');
        const tNorm = normalizar(tipoStr);
        const esRV = (tNorm === 'Cedear' || tNorm === 'Acciones');
        const esRF = !esRV && !esCash;

        if (mov.includes('aporte')) {
            if (esCash) {
                cfTOTAL.push({ d: fechaObj, v: -montoUSD });
                aporteNetoUSD += montoUSD;
                bridgeAportes += montoUSD;
            }
        } else if (mov.includes('retiro')) {
            if (esCash) {
                cfTOTAL.push({ d: fechaObj, v: montoUSD });
                aporteNetoUSD -= montoUSD;
                bridgeRetiros += montoUSD;
            }
        }

        if (mov.includes('dividen')) bridgeDividendos += montoUSD;
        if (mov.includes('renta') || mov.includes('interes')) bridgeRenta += montoUSD;
        if (mov.includes('amortiz')) bridgeAmortizaciones += montoUSD;

        if (esRV) {
            if (mov.includes('compra') || mov.includes('suscrip')) {
                cfRV.push({ d: fechaObj, v: -montoUSD });
                comprasNetasRV += montoUSD;
            } else if (mov.includes('venta') || mov.includes('rescate') || mov.includes('dividen')) {
                cfRV.push({ d: fechaObj, v: montoUSD });
                comprasNetasRV -= montoUSD;
            }
        } else if (esRF) {
            if (mov.includes('compra') || mov.includes('suscrip')) {
                cfRF.push({ d: fechaObj, v: -montoUSD });
                comprasNetasRF += montoUSD;
            } else if (mov.includes('venta') || mov.includes('rescate') || mov.includes('renta') || mov.includes('interes') || mov.includes('amortiz')) {
                cfRF.push({ d: fechaObj, v: montoUSD });
                comprasNetasRF -= montoUSD;
            }
        }
    });

    if (valFinalTotal > 0) cfTOTAL.push({ d: hoy, v: valFinalTotal });
    if (valFinalRV > 0) cfRV.push({ d: hoy, v: valFinalRV });
    if (valFinalRF > 0) cfRF.push({ d: hoy, v: valFinalRF });

    let gananciaNetaTotalUSD = valFinalTotal - valInicialTotal - aporteNetoUSD;
    let baseInvertidaTotal = valInicialTotal + aporteNetoUSD;
    let retAbsolutoTotal = baseInvertidaTotal > 0 ? (gananciaNetaTotalUSD / baseInvertidaTotal) : 0;

    let gananciaNetaRVUSD = valFinalRV - valInicialRV - comprasNetasRV;
    let baseInvertidaRV = valInicialRV + Math.max(0, comprasNetasRV);
    let retAbsolutoRV = baseInvertidaRV > 0 ? (gananciaNetaRVUSD / baseInvertidaRV) : 0;

    let gananciaNetaRFUSD = valFinalRF - valInicialRF - comprasNetasRF;
    let baseInvertidaRF = valInicialRF + Math.max(0, comprasNetasRF);
    let retAbsolutoRF = baseInvertidaRF > 0 ? (gananciaNetaRFUSD / baseInvertidaRF) : 0;
    let plValuacionUSD = valFinalTotal - valInicialTotal - bridgeAportes + bridgeRetiros - bridgeDividendos - bridgeRenta - bridgeAmortizaciones;

    return {
        anio: anioActual,
        fechaCorteStr: fechaCorteStr,
        diasPeriodo: Math.max(1, Math.round((hoy.getTime() - FECHA_CORTE.getTime()) / 86400000)),
        tirTotal: calcXIRR(cfTOTAL),
        tirVariable: calcXIRR(cfRV),
        tirFija: calcXIRR(cfRF),
        aporteNetoUSD: aporteNetoUSD,
        retAbsolutoTotal: retAbsolutoTotal,
        retAbsolutoVariable: retAbsolutoRV,
        retAbsolutoFija: retAbsolutoRF,
        gananciaNetaTotalUSD: gananciaNetaTotalUSD,
        saldoInicialUSD: valInicialTotal,
        saldoFinalUSD: valFinalTotal,
        aportesUSD: bridgeAportes,
        retirosUSD: bridgeRetiros,
        dividendosUSD: bridgeDividendos,
        rentaUSD: bridgeRenta,
        amortizacionesUSD: bridgeAmortizaciones,
        plValuacionUSD: plValuacionUSD
    };
}

function calcXIRR(values, guess = 0.1) {
    if (!values || values.length < 2) return 0;
    let hasPos = false, hasNeg = false;
    for (let x of values) { if (x.v > 0) hasPos = true; if (x.v < 0) hasNeg = true; }
    if (!hasPos || !hasNeg) return 0;
    values.sort((a, b) => a.d - b.d);

    const t0 = values[0].d.getTime();
    const van = (r) => {
        let suma = 0;
        for (let j = 0; j < values.length; j++) {
            const dt = (values[j].d.getTime() - t0) / 31536000000.0;
            const div = Math.pow(1.0 + r, dt);
            if (!div || !isFinite(div)) return NaN;
            suma += values[j].v / div;
        }
        return suma;
    };

    const runNewton = (g) => {
        let x0 = g, x1 = 0.0, tol = 1e-5, maxIter = 50;
        for (let i = 0; i < maxIter; i++) {
            if (x0 <= -1) x0 = -0.99999999;
            let fv = 0.0, fd = 0.0;
            for (let j = 0; j < values.length; j++) {
                let t = values[j].d.getTime();
                let dt = (t - t0) / 31536000000.0;
                if (dt < 0.00001 && j > 0) dt = 0.00001;
                let div = Math.pow(1.0 + x0, dt);
                if (div === 0 || !isFinite(div)) return null;
                fv += values[j].v / div;
                fd -= (dt * values[j].v) / (div * (1.0 + x0));
            }
            if (Math.abs(fd) < 1e-9) return null;
            x1 = x0 - fv / fd;
            if (Math.abs(x1 - x0) <= tol) return x1;
            x0 = x1;
        }
        return null;
    };

    let resultadoNewton = null;
    for (let g of [0.1, -0.5, 0.9, -0.9, 2.0, 0.5, -0.1, -0.2, -0.3, -0.6, -0.7]) {
        let result = runNewton(g);
        if (result !== null && Math.abs(result) < 100 && isFinite(result) && result > -0.999) {
            resultadoNewton = result;
            break;
        }
    }

    if (resultadoNewton !== null) return resultadoNewton;

    let lo = -0.99, hi = 5.0;
    let vanLo = van(lo), vanHi = van(hi);

    if (isNaN(vanLo) || isNaN(vanHi) || (vanLo > 0 && vanHi > 0) || (vanLo < 0 && vanHi < 0)) {
        hi = 50.0;
        vanHi = van(hi);
        if (isNaN(vanHi) || (vanLo > 0 && vanHi > 0) || (vanLo < 0 && vanHi < 0)) {
            return 0;
        }
    }

    let mid = 0, vanMid = 0;
    const maxIterBiseccion = 100;
    const tolBiseccion = 1e-6;

    for (let i = 0; i < maxIterBiseccion; i++) {
        mid = (lo + hi) / 2;
        vanMid = van(mid);
        if (isNaN(vanMid)) { hi = mid; continue; }
        if (Math.abs(vanMid) < tolBiseccion || (hi - lo) / 2 < tolBiseccion) return mid;
        if ((vanLo > 0 && vanMid > 0) || (vanLo < 0 && vanMid < 0)) {
            lo = mid; vanLo = vanMid;
        } else {
            hi = mid; vanHi = vanMid;
        }
    }
    return mid;
}

function calcCostoPromedioPonderado(costoEvents, anio, hoy) {
    const inicioAnio = new Date(anio, 0, 1);
    const finAnio = new Date(anio, 11, 31);
    const corte = (hoy < finAnio) ? hoy : finAnio;
    const diasTranscurridos = Math.floor((corte - inicioAnio) / 86400000) + 1;
    if (diasTranscurridos <= 0 || !costoEvents || costoEvents.length === 0) {
        return { costoProm: 0, dias: 0, esParcial: false };
    }

    let costoActual = 0;
    for (let i = 0; i < costoEvents.length; i++) {
        if (costoEvents[i].d <= inicioAnio) costoActual = costoEvents[i].c;
        else break;
    }

    let puntos = [{ d: inicioAnio, c: costoActual }];
    costoEvents.forEach(e => {
        if (e.d > inicioAnio && e.d <= corte) puntos.push({ d: e.d, c: e.c });
    });
    puntos.push({ d: new Date(corte.getTime() + 86400000), c: null });

    let acumulado = 0;
    for (let i = 0; i < puntos.length - 1; i++) {
        let dias = Math.floor((puntos[i + 1].d - puntos[i].d) / 86400000);
        acumulado += puntos[i].c * dias;
    }

    return {
        costoProm: acumulado / diasTranscurridos,
        dias: diasTranscurridos,
        esParcial: corte < finAnio
    };
}

function cleanNum(v) {
    if (typeof v === 'number') return v;
    if (!v) return 0;
    let s = String(v).trim().replace('%', '');
    if (s.includes('.') && s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.includes(',')) s = s.replace(',', '.');
    return parseFloat(s) || 0;
}

function normalizar(t) {
    const original = t;
    t = (t || '').toLowerCase();
    if (t.includes('cedear')) return 'Cedear';
    if (t.includes('accion')) return 'Acciones';
    if (t.includes('bono')) return 'Bonos';
    if (t.includes('negociable') || t.includes(' on')) return 'ONs';
    if (t.includes('fci')) return 'FCI';
    if (t.includes('cartera')) return 'Cartera';
    if (t.includes('liquidez')) return 'Liquidez';
    if (original && original.trim() !== '' && original.trim() !== 'undefined') {
        console.warn('[normalizar] Tipo no reconocido: "' + original + '" → clasificado como Otros');
    }
    return 'Otros';
}

function top(l, t1, t2) {
    return l.filter(x => x.tipo === t1 || x.tipo === t2)
            .sort((a, b) => b.valorUSD - a.valorUSD)
            .slice(0, 10);
}

function calcSortino(valores) {
    if (!valores || valores.length < 30) return 0;
    const r = [];
    for (let i = 1; i < valores.length; i++) {
        if (valores[i - 1] > 0) r.push((valores[i] - valores[i - 1]) / valores[i - 1]);
    }
    if (r.length === 0) return 0;
    const MAR_ANUAL = 0.05;
    const MAR_DIARIO = MAR_ANUAL / 252;
    const mean = r.reduce((a, b) => a + b, 0) / r.length;
    const downside = r.filter(x => x < MAR_DIARIO);
    if (downside.length === 0) return 10;
    const sumSqDown = downside.reduce((a, b) => a + Math.pow(b - MAR_DIARIO, 2), 0);
    const downDev = Math.sqrt(sumSqDown / r.length);
    if (downDev === 0) return 10;
    return ((mean - MAR_DIARIO) * 252) / (downDev * Math.sqrt(252));
}

function calcDGR(stats) {
    const yrs = Object.keys(stats.years).map(y => parseInt(y)).sort((a, b) => b - a);
    const currentYear = new Date().getFullYear();
    const completedYears = yrs.filter(y => y < currentYear);
    if (completedYears.length < 2) return 0;
    const lastFullY = stats.years[completedYears[0]] || 0;
    const prevFullY = stats.years[completedYears[1]] || 0;
    if (prevFullY === 0) return 1;
    return (lastFullY / prevFullY) - 1;
}

function calcNextPay(lista) {
    if (!lista || lista.length === 0) return { tk: '---', days: 999, amt: 0, renta: 0, capital: 0, count: 0, tickers: [] };
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);

    let proximaFecha = null;
    for (let f of lista) {
        let d = new Date(f.fecha);
        d.setHours(0, 0, 0, 0);
        if (d >= hoy) { proximaFecha = d; break; }
    }
    if (!proximaFecha) return { tk: '---', days: 999, amt: 0, renta: 0, capital: 0, count: 0, tickers: [] };

    const proximaStr = proximaFecha.toISOString().slice(0, 10);
    let pagosDelDia = lista.filter(f => {
        let d = new Date(f.fecha);
        d.setHours(0, 0, 0, 0);
        return d.toISOString().slice(0, 10) === proximaStr;
    });

    let totalRenta = 0, totalCapital = 0, tickers = [];
    pagosDelDia.forEach(p => {
        totalRenta += p.renta;
        totalCapital += p.capital;
        tickers.push(p.ticker);
    });

    let diff = Math.ceil((proximaFecha - hoy) / (1000 * 60 * 60 * 24));

    return {
        tk: tickers.join(', '),
        days: diff,
        renta: totalRenta,
        capital: totalCapital,
        amt: totalRenta + totalCapital,
        count: tickers.length,
        tickers: tickers
    };
}

function leerQuant(h) {
    if (!h) return { volatilidadAnualizada: 0, sharpeRatio: 0, betaCartera: 0, maxDrawdown: 0 };
    const v = h.getRange("H2:H6").getValues();
    return { volatilidadAnualizada: v[1][0], sharpeRatio: v[2][0], betaCartera: v[3][0], maxDrawdown: v[4][0] };
}

function leerEvo(h) {
    if (!h) return { fechas: [], valores: [], benchmark: [], cedears: [] };
    const lastRow = h.getLastRow();
    if (lastRow < 2) return { fechas: [], valores: [], benchmark: [], cedears: [] };
    const numCols = Math.min(4, Math.max(3, h.getLastColumn()));
    const colLetter = numCols >= 4 ? "D" : "C";
    const d = h.getRange("A2:" + colLetter + lastRow).getValues();
    return {
        fechas: d.map(x => x[0]),
        valores: d.map(x => cleanNum(x[1])),
        benchmark: d.map(x => cleanNum(x[2])),
        cedears: d.map(x => numCols >= 4 ? cleanNum(x[3]) : 0)
    };
}

function leerProyecciones(h) {
    if (!h) return { porFecha: [] };
    const d = h.getRange("A2:E" + h.getLastRow()).getValues();
    return { porFecha: d.filter(r => r[4] > 0).map(r => ({ ticker: r[0], fecha: r[1], renta: r[2], capital: r[3] })) };
}

function submuestrearHistPrecios(rawMap, maxPoints) {
    const MAX_POINTS = maxPoints || 150;
    const finalMap = {};
    Object.keys(rawMap).forEach(tk => {
        let points = rawMap[tk];
        if (points.length <= MAX_POINTS) {
            finalMap[tk] = points;
        } else {
            const reduced = [];
            const step = Math.ceil(points.length / MAX_POINTS);
            for (let j = 0; j < points.length; j += step) reduced.push(points[j]);
            const lastReal = points[points.length - 1];
            const lastSaved = reduced[reduced.length - 1];
            if (lastReal[0] !== lastSaved[0]) reduced.push(lastReal);
            finalMap[tk] = reduced;
        }
    });
    return finalMap;
}

function leerHistPreciosCompleto(h) {
    if (!h) return {};
    const d = h.getRange("A2:C" + h.getLastRow()).getValues();
    const rawMap = {};
    for (let i = 0; i < d.length; i++) {
        const r = d[i];
        if (r[1] && r[2] > 0 && r[0] instanceof Date) {
            const tk = String(r[1]).toUpperCase().trim();
            if (!rawMap[tk]) rawMap[tk] = [];
            rawMap[tk].push([r[0].getTime(), r[2]]);
        }
    }
    Object.keys(rawMap).forEach(tk => rawMap[tk].sort((a, b) => a[0] - b[0]));
    return rawMap;
}

function obtenerValorAyer(hoja) {
    if (!hoja) return 0;
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return 0;
    return parseFloat(hoja.getRange(lastRow, 2).getValue());
}

function procesarFlujosAcumulativosMensuales(hoja) {
    if (!hoja) return { labels: [], monthly: [], cumulative: [] };
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { labels: [], monthly: [], cumulative: [] };
    const rango = hoja.getRange("A2:E" + ultimaFila).getValues();
    const flujosMensuales = {};
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    rango.forEach(fila => {
        const fecha = new Date(fila[1]);
        const montoTotal = (fila[2] || 0) + (fila[3] || 0);
        if (fecha instanceof Date && !isNaN(fecha) && fecha >= hoy && montoTotal > 0) {
            const mesClave = `${fecha.getFullYear()}-${String(fecha.getMonth() + 1).padStart(2, '0')}`;
            flujosMensuales[mesClave] = (flujosMensuales[mesClave] || 0) + montoTotal;
        }
    });
    const mesesOrdenados = Object.keys(flujosMensuales).sort();
    const labels = [], monthlyData = [], cumulativeData = [];
    let acumulado = 0;
    mesesOrdenados.forEach(mesClave => {
        const [ano, mes] = mesClave.split('-');
        const nombreMes = new Date(ano, mes - 1).toLocaleString('es-AR', { month: 'short' });
        labels.push(`${nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1)}/${ano.slice(2)}`);
        const montoMes = flujosMensuales[mesClave];
        monthlyData.push(montoMes);
        acumulado += montoMes;
        cumulativeData.push(acumulado);
    });
    return { labels, monthly: monthlyData, cumulative: cumulativeData };
}

function procesarFlujosSemestrales(hoja) {
    if (!hoja) return { labels: [], datasets: [] };
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { labels: [], datasets: [] };
    const rango = hoja.getRange("A2:E" + ultimaFila).getValues();
    const datosAgregados = {};
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    rango.forEach(fila => {
        const ticker = fila[0];
        const fecha = new Date(fila[1]);
        const montoTotal = (fila[2] || 0) + (fila[3] || 0);
        if (ticker && fecha instanceof Date && !isNaN(fecha) && fecha >= hoy && montoTotal > 0) {
            const ano = fecha.getFullYear().toString().slice(-2);
            const semestreLabel = fecha.getMonth() < 6 ? `Ene/${ano}` : `Jul/${ano}`;
            if (!datosAgregados[semestreLabel]) datosAgregados[semestreLabel] = {};
            datosAgregados[semestreLabel][ticker] = (datosAgregados[semestreLabel][ticker] || 0) + montoTotal;
        }
    });
    const labels = Object.keys(datosAgregados).sort((a, b) => {
        const [m1, y1] = a.split('/'); const [m2, y2] = b.split('/');
        if (y1 !== y2) return parseInt(y1) - parseInt(y2);
        return m1 === 'Ene' ? -1 : 1;
    });
    const tickersUnicos = [...new Set(rango.map(fila => fila[0]).filter(Boolean))];
    const datasets = tickersUnicos.map(ticker => {
        const data = labels.map(label => datosAgregados[label][ticker] || 0);
        return { label: ticker, data: data };
    }).filter(ds => ds.data.some(d => d > 0));
    return { labels, datasets };
}

// =================================================================================
// PASO 3.1 de la consolidación (ver debate "Portafolio Ultimate"): esta función
// ya NO reimplementa el loop de posiciones. Delega en snapshotAFecha(), de
// MotorPosiciones.gs, que aplica las mismas reglas que el resto del motor
// (Método B de amortización, tope de venta a cantidad tenida, venta sin stock
// previo = ganancia total). Antes de este paso, esta función tenía su propia
// versión simplificada de esas reglas, SIN el tope de venta ni la regla de
// venta-sin-stock — validado en paralelo contra el sistema viejo el 06/09/2026
// (VALIDACION_MOTOR_TEMP): coincide en todos los tickers salvo BRKB y HMY,
// que difieren a propósito por esas dos reglas defensivas.
//
// actualizarUniversoTickers() (más abajo) y Bombonera.gs siguen recibiendo
// exactamente el mismo formato de salida { ticker: {q, costo, tipo} } que
// siempre — no requirieron ningún cambio.
// =================================================================================
function calcularCantidadesNetas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logData = ss.getSheetByName(HOJAS.LOG).getDataRange().getValues();
    const logSinHeader = logData.slice(1);
    logSinHeader.sort((a, b) => new Date(a[1]) - new Date(b[1]));

    const resultado = snapshotAFecha(logSinHeader, HOY_SIMULADA);

    let cantidades = {};
    Object.keys(resultado.posiciones).forEach(tk => {
        const p = resultado.posiciones[tk];
        cantidades[tk] = { q: p.q, costo: p.costo, tipo: p.tipo };
    });

    if (resultado.advertencias.length > 0) {
        resultado.advertencias.forEach(a => console.warn('[calcularCantidadesNetas] ' + a.mensaje));
    }

    return cantidades;
}

function actualizarUniversoTickers() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCartera = ss.getSheetByName(HOJAS.CARTERA);
    if (!hojaCartera) { console.error("No se encontró 'Cartera'"); return; }

    const cantidades = calcularCantidadesNetas();

    const ordenTipos = ["Cedear", "Acciones", "Bonos", "O Negociable", "FCI", "Cartera admin", "Liquidez"];

    let porTipo = {};
    ordenTipos.forEach(t => porTipo[t] = []);

    Object.keys(cantidades).forEach(tk => {
        const item = cantidades[tk];
        if (item.q <= 0.01) return;

        const tNorm = normalizar(item.tipo);
        let etiqueta = null;
        if (tNorm === 'Cedear') etiqueta = 'Cedear';
        else if (tNorm === 'Acciones') etiqueta = 'Acciones';
        else if (tNorm === 'Bonos') etiqueta = 'Bonos';
        else if (tNorm === 'ONs') etiqueta = 'O Negociable';
        else if (tNorm === 'FCI') etiqueta = 'FCI';
        else if (tNorm === 'Cartera') etiqueta = 'Cartera admin';
        else etiqueta = 'Liquidez';

        if (porTipo[etiqueta]) porTipo[etiqueta].push({ ticker: tk, q: item.q, costo: item.costo });
    });

    let filasWX = [];
    let filasY = [];
    let filasAC = [];

    ordenTipos.forEach(tipo => {
        const itemsOrdenados = porTipo[tipo].sort((a, b) => a.ticker.localeCompare(b.ticker));
        itemsOrdenados.forEach(item => {
            filasWX.push([item.ticker, tipo]);
            filasY.push([item.q]);
            filasAC.push([item.costo]);
        });
    });

    const COL_W = 23, COL_X = 24, COL_Y = 25, COL_AC = 29;
    const ultimaFilaVieja = hojaCartera.getLastRow();
    if (ultimaFilaVieja >= 2) {
        hojaCartera.getRange(2, COL_W, Math.max(ultimaFilaVieja - 1, 1), 2).clearContent();
        hojaCartera.getRange(2, COL_Y, Math.max(ultimaFilaVieja - 1, 1), 1).clearContent();
        hojaCartera.getRange(2, COL_AC, Math.max(ultimaFilaVieja - 1, 1), 1).clearContent();
    }

    if (filasWX.length > 0) {
        hojaCartera.getRange(2, COL_W, filasWX.length, 2).setValues(filasWX);
        hojaCartera.getRange(2, COL_Y, filasY.length, 1).setValues(filasY);
        hojaCartera.getRange(2, COL_AC, filasAC.length, 1).setValues(filasAC);
        
        hojaCartera.getRange(2, COL_Y, filasY.length, 1).setNumberFormat("#,##0.00");
        hojaCartera.getRange(2, COL_AC, filasAC.length, 1).setNumberFormat("#,##0.00");
    }

    const FILA_USD = 90;
    const ccl = obtenerCCLActual();
    const cajaVirtual = calcularCajaVirtual();

    hojaCartera.getRange(FILA_USD, COL_W, 1, 2).setValues([["USD", "Liquidez"]]);
    hojaCartera.getRange(FILA_USD, 25, 1, 5).setValues([[
        cajaVirtual,
        ccl,
        cajaVirtual * ccl,
        cajaVirtual,
        cajaVirtual
    ]]);

    console.log(`✅ Universo de tickers, Cantidades y Costos actualizados: ${filasWX.length} tickers vivos | ${Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "dd/MM/yyyy HH:mm:ss")}`);
}

function obtenerCCLActual() {
    let ccl = 1000;
    try {
        const val = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.PRECIOS).getRange("B4").getValue();
        if (typeof val === 'number' && val > 500) ccl = val;
    } catch (e) { }
    return ccl;
}

// =================================================================================
// PASO 3.2 de la consolidación (ver debate "Portafolio Ultimate"): esta función
// ya NO reimplementa el loop de cashflows/cantidades. Los cashflows por ticker
// salen de ledgerCompleto() (MotorPosiciones.gs), que ya aplica Método B, el
// tope de venta y la regla de venta-sin-stock. La cantidad actual (para saber
// si la posición sigue abierta y corresponde sumar el valor de hoy al XIRR)
// sale de snapshotAFecha() con corte a HOY_SIMULADA — la misma fuente que ya
// usa calcularCantidadesNetas() desde el Paso 3.1.
// =================================================================================
function actualizarXirrCartera() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCartera = ss.getSheetByName(HOJAS.CARTERA);
    if (!hojaCartera) { console.error("No se encontró 'Cartera'"); return; }

    const logData = ss.getSheetByName(HOJAS.LOG).getDataRange().getValues();
    const logSinHeader = logData.slice(1);
    logSinHeader.sort((a, b) => new Date(a[1]) - new Date(b[1]));

    const ledger = ledgerCompleto(logSinHeader);
    const snapshotHoy = snapshotAFecha(logSinHeader, HOY_SIMULADA);

    let cashflowsPorTicker = {};
    ledger.transacciones.forEach(t => {
        if (!t.cashflow) return; // splits y filas salteadas no aportan cashflow
        if (!cashflowsPorTicker[t.ticker]) cashflowsPorTicker[t.ticker] = [];
        cashflowsPorTicker[t.ticker].push({ d: t.fecha, v: t.cashflow });
    });

    const columnaA = hojaCartera.getRange("A2:A" + hojaCartera.getMaxRows()).getValues();
    let ultimaFilaConTicker = 0;
    for (let i = 0; i < columnaA.length; i++) {
        if (String(columnaA[i][0]).trim() !== "") ultimaFilaConTicker = i + 1;
    }
    if (ultimaFilaConTicker === 0) { console.warn("Cartera vacía (columna A)"); return; }

    const dataCartera = hojaCartera.getRange(2, 1, ultimaFilaConTicker, 6).getValues();
    let resultados = [];

    dataCartera.forEach(row => {
        const ticker = String(row[0]).toUpperCase().trim();
        const valorActualUSD = cleanNum(row[5]);

        if (!ticker || !cashflowsPorTicker[ticker]) { resultados.push(0); return; }

        const qActual = (snapshotHoy.posiciones[ticker] || {}).q || 0;
        const cfs = [...cashflowsPorTicker[ticker]];
        if (qActual > 0.01 && valorActualUSD > 0) {
            cfs.push({ d: new Date(), v: valorActualUSD });
        }

        const xirr = calcXIRR(cfs);
        resultados.push(xirr);
    });

    hojaCartera.getRange(2, 9, resultados.length, 1).setValues(resultados.map(x => [x]));
    hojaCartera.getRange(2, 9, resultados.length, 1).setNumberFormat("0.00%");

    console.log(`✅ XIRR actualizado en Cartera para ${resultados.length} tickers.`);
}

function calcularCajaVirtual(logSinHeaderExterno) {
    let logSinHeader;
    if (logSinHeaderExterno) {
        logSinHeader = logSinHeaderExterno;
    } else {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const logData = ss.getSheetByName(HOJAS.LOG).getDataRange().getValues();
        logSinHeader = logData.slice(1);
    }

    let cajaVirtual = SALDO_INICIAL_CAJA;

    logSinHeader.forEach(row => {
        const fecha = new Date(row[1]);
        if (isNaN(fecha.getTime()) || fecha < FECHA_CORTE_CAJA) return;

        const ticker = String(row[3]).toUpperCase().trim();
        const mov = String(row[5]).toLowerCase().trim();
        const montoUSD = Math.abs(cleanNum(row[10]));
        const esCash = (ticker === 'USD' || ticker === 'CASH');

        if (mov.includes('compra') || mov.includes('suscripcion') || mov.includes('retiro') || mov.includes('canje_entrada')) {
            cajaVirtual -= montoUSD;
        } else if (mov.includes('venta') || mov.includes('rescate') ||
                   mov.includes('dividendo') || mov.includes('renta') || mov.includes('interes') ||
                   mov.includes('amortiza') || mov.includes('canje_salida')) {
            cajaVirtual += montoUSD;
        } else if (mov.includes('aporte')) {
            // "Aporte" es ambiguo: puede ser plata nueva entrando a la cuenta
            // (ticker=USD/CASH -> suma a la caja) o una suscripción a un fondo
            // pagada con esa plata (cualquier otro ticker -> resta, como una compra).
            if (esCash) cajaVirtual += montoUSD;
            else cajaVirtual -= montoUSD;
        }
    });

    return cajaVirtual;
}

// Caja reconstruida desde el saldo contable confiable del 01/02/2026.
function calcularCajaVirtualHasta(logSinHeader, fechaHasta) {
    const hastaMs = fechaHasta.getTime();
    return calcularCajaVirtual((logSinHeader || []).filter(row => {
        const fecha = new Date(row[1]);
        return !isNaN(fecha.getTime()) && fecha.getTime() <= hastaMs;
    }));
}
