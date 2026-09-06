// =================================================================================
// ===  TITANIUM v2 — PRECIOS_AUTOMATICOS.GS                                    ===
// ===  Tres actualizadores: CCL, Compass/CAFCI, data912                        ===
// =================================================================================

// =================================================================================
// 1. CCL — dolarapi.com, escribe en Precios!Q:R + Precios!B4
// =================================================================================
function REGISTRAR_CCL_INTELIGENTE() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPrecios = ss.getSheetByName(HOJAS.PRECIOS);
    const tz = ss.getSpreadsheetTimeZone();
    const ahora = new Date();

    const diaSemana = parseInt(Utilities.formatDate(ahora, tz, "u"));
    const horaActual = parseInt(Utilities.formatDate(ahora, tz, "H"));
    if (diaSemana > 5 || horaActual < 11 || horaActual >= 19) {
        console.log("Fuera de horario o fin de semana.");
        return;
    }

    if (!hojaPrecios) { console.error("No se encontró 'Precios'"); return; }

    try {
        const response = UrlFetchApp.fetch("https://dolarapi.com/v1/dolares/contadoconliqui");
        const data = JSON.parse(response.getContentText());

        const hoyTextoISO = Utilities.formatDate(ahora, tz, "yyyy-MM-dd");
        const fechaApiISO = data.fechaActualizacion.split('T')[0];

        if (hoyTextoISO !== fechaApiISO) {
            console.warn("Dato de API viejo (posible feriado). No se escribe nada.");
            return;
        }

        const precio = parseFloat(data.venta);

        const COL_FECHA = 17, COL_PRECIO = 18;
        const ultimaFilaConDatos = hojaPrecios.getRange(1, COL_FECHA, hojaPrecios.getMaxRows(), 1)
            .getValues()
            .reduce((last, val, idx) => (val[0] !== "" && val[0] !== null) ? idx + 1 : last, 1);

        let filaDestino = ultimaFilaConDatos + 1;

        if (ultimaFilaConDatos >= 2) {
            const fechaHoyVisual = Utilities.formatDate(ahora, tz, "d/M/yyyy");
            const valoresQ = hojaPrecios.getRange(2, COL_FECHA, ultimaFilaConDatos - 1, 1).getDisplayValues();
            for (let i = 0; i < valoresQ.length; i++) {
                if (valoresQ[i][0] === fechaHoyVisual || valoresQ[i][0] === hoyTextoISO.split('-').reverse().join('/')) {
                    filaDestino = i + 2;
                    break;
                }
            }
        }

        const celdaFecha = hojaPrecios.getRange(filaDestino, COL_FECHA);
        const celdaPrecio = hojaPrecios.getRange(filaDestino, COL_PRECIO);

        celdaFecha.setValue(hoyTextoISO);
        celdaPrecio.setValue(precio);
        celdaFecha.setNumberFormat("dd/mm/yyyy");

        if (filaDestino > 2) {
            const formatoPrecioArriba = hojaPrecios.getRange(filaDestino - 1, COL_PRECIO).getNumberFormat();
            celdaPrecio.setNumberFormat(formatoPrecioArriba);
        }

        hojaPrecios.getRange("B4").setValue(precio);

        SpreadsheetApp.flush();
        console.log(`Éxito. Fila ${filaDestino} (Q:R) y B4 actualizados con $${precio}.`);

    } catch (e) {
        console.error("Error CCL: " + e.message);
    }
}

// =================================================================================
// 2. COMPASS / CAFCI — fondos en bloque protegido (filas 1-19)
// =================================================================================
function ACTUALIZAR_PRECIOS_CAFCI() {
    const HOJA_PRECIOS = 'Precios';

    const hoy = new Date();
    const diaSemana = hoy.getDay();
    if (diaSemana === 0 || diaSemana === 6) {
        Logger.log('Fin de semana detectado. No se actualizan fondos.');
        return;
    }

    const FONDOS = [
        { ticker: 'CAU$D',    codigoCAFCI: 2112, usarNativoUSD: false },
        { ticker: 'CBIDEA',   codigoCAFCI: 2106, usarNativoUSD: false },
        { ticker: 'CCREC2',   codigoCAFCI: 788,  usarNativoUSD: false },
        { ticker: 'COPPORT',  codigoCAFCI: 338,  usarNativoUSD: false },
        { ticker: 'COCORMA',  codigoCAFCI: 2517, usarNativoUSD: false },
        { ticker: 'COCOSPPA', codigoCAFCI: 5496, usarNativoUSD: false },
        { ticker: 'ALLFD',    codigoCAFCI: 5901, usarNativoUSD: true  },
        { ticker: 'BAHUSD',   codigoCAFCI: 2096, usarNativoUSD: true  },
    ];

    Logger.log('Descargando planilla diaria CAFCI...');
    let blob;
    try {
        const resp = UrlFetchApp.fetch('https://download.cafci.org.ar/Planilla_Diaria_A.xlsx', {
            muteHttpExceptions: true,
            headers: { 'User-Agent': 'Mozilla/5.0', 'Referer': 'https://www.cafci.org.ar/' }
        });
        if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
        blob = resp.getBlob();
    } catch (e) {
        Logger.log('Error de descarga: ' + e.message + '. Se conservan los precios anteriores.');
        return;
    }

    let tempSheet;
    try {
        blob.setName('planilla_cafci.xlsx');
        const tempFile = Drive.Files.create(
            { name: 'TEMP_CAFCI_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS },
            blob
        );
        tempSheet = SpreadsheetApp.openById(tempFile.id);
    } catch (e) {
        Logger.log('Error de conversión API Drive: ' + e.message);
        return;
    }

    let preciosNuevos = {};
    try {
        const datosExcel = tempSheet.getSheets()[0].getDataRange().getValues();

        const indicePorCodigo = {};
        for (let r = 10; r < datosExcel.length; r++) {
            const codigo = datosExcel[r][20];
            if (typeof codigo === 'number') indicePorCodigo[codigo] = r;
        }

        for (const fondo of FONDOS) {
            const fila = indicePorCodigo[fondo.codigoCAFCI];
            if (fila === undefined) { Logger.log(`⚠️ No se encontró el código CAFCI ${fondo.codigoCAFCI} (${fondo.ticker}) hoy.`); continue; }
            const row = datosExcel[fila];
            const valorNativoRaw = row[5];
            const valorReexpresadoRaw = row[8];

            let valorFinalRaw = valorNativoRaw;
            if (!fondo.usarNativoUSD && typeof valorReexpresadoRaw === 'number' && valorReexpresadoRaw > 0) {
                valorFinalRaw = valorReexpresadoRaw;
            }
            if (typeof valorFinalRaw === 'number') {
                preciosNuevos[fondo.ticker] = valorFinalRaw / 1000;
            }
        }
    } catch (e) {
        Logger.log('Error al procesar Excel: ' + e.message);
    } finally {
        try { DriveApp.getFileById(tempSheet.getId()).setTrashed(true); } catch (e) { }
    }

    if (Object.keys(preciosNuevos).length === 0) {
        Logger.log('No se pudieron extraer precios. Se conservan los anteriores.');
        return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPrecios = ss.getSheetByName(HOJA_PRECIOS);
    if (!hojaPrecios) { Logger.log('No se encontró la pestaña ' + HOJA_PRECIOS); return; }

    const datosPrecios = hojaPrecios.getDataRange().getValues();
    for (let i = 1; i < datosPrecios.length; i++) {
        let tickerActual = String(datosPrecios[i][0]).toUpperCase().trim();
        if (preciosNuevos[tickerActual]) {
            hojaPrecios.getRange(i + 1, 2).setValue(preciosNuevos[tickerActual]);
            Logger.log(`✅ Actualizado: ${tickerActual} -> ${preciosNuevos[tickerActual]}`);
        }
    }

    Logger.log('¡Actualización de Fondos CAFCI finalizada con éxito!');
}

// =================================================================================
// 3. DATA912 — Cedears / Acciones / Bonos / ONs, bloque dinámico desde fila 20
//    FIX: antes, si data912 respondía 200 pero devolvía precio=0 para un ticker
//    puntual (típico en pre-market, mercado cerrado), ese cero pisaba el precio
//    bueno anterior sin ningún control — el respaldo solo actuaba si el endpoint
//    entero fallaba, no ticker por ticker. Ahora se preserva el precio anterior
//    de CADA ticker individual cuando el nuevo dato es 0/ausente.
// =================================================================================
function actualizarPrecios() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPrecios = ss.getSheetByName(HOJAS.PRECIOS);
    const hojaCartera = ss.getSheetByName(HOJAS.CARTERA);

    if (!hojaPrecios || !hojaCartera) {
        console.error("No se encontró 'Precios' o 'Cartera'");
        return;
    }

    const lastRowCartera = hojaCartera.getLastRow();
    if (lastRowCartera < 2) { console.warn("Cartera vacía"); return; }

    const dataCartera = hojaCartera.getRange(2, 1, lastRowCartera - 1, 2).getValues();

    const grupos = {
        "Cedear":       new Set(),
        "Acciones":     new Set(),
        "Bonos":        new Set(),
        "O Negociable": new Set()
    };
    dataCartera.forEach(([ticker, tipo]) => {
        const t = String(ticker).trim().toUpperCase();
        const tp = String(tipo).trim();
        if (grupos[tp] && t) grupos[tp].add(t);
    });

    const endpoints = {
        "Cedear":       "https://data912.com/live/arg_cedears",
        "Acciones":     "https://data912.com/live/arg_stocks",
        "Bonos":        "https://data912.com/live/arg_bonds",
        "O Negociable": "https://data912.com/live/arg_corp"
    };

    const precios = {};
    const fetchOk = {};

    Object.entries(endpoints).forEach(([tipo, url]) => {
        fetchOk[tipo] = false;
        try {
            const resp = UrlFetchApp.fetch(url, {
                muteHttpExceptions: true,
                headers: { "User-Agent": "Mozilla/5.0" }
            });
            if (resp.getResponseCode() !== 200) {
                console.error(`HTTP ${resp.getResponseCode()} → ${url}. Se conserva el bloque anterior de ${tipo}.`);
                return;
            }
            JSON.parse(resp.getContentText()).forEach(item => {
                const sym = String(item.symbol || "").trim().toUpperCase();
                precios[sym] = item.c || 0;
            });
            fetchOk[tipo] = true;
        } catch (e) {
            console.error(`Error en ${url}: ${e.message}. Se conserva el bloque anterior de ${tipo}.`);
        }
    });

    if (!Object.values(fetchOk).some(v => v)) {
        console.warn("Ningún endpoint respondió. No se modifica Precios.");
        return;
    }

    const FILA_INICIO = 20;
    const ultimaFilaActual = hojaPrecios.getLastRow();

    const ordenCategorias = ["Cedear", "Acciones", "Bonos", "O Negociable"];
    const bloquesViejosPorCategoria = {};
    ordenCategorias.forEach(c => bloquesViejosPorCategoria[c] = []);

    // NUEVO: mapa ticker -> {precio, numberFormat}, para poder recuperar el
    // último precio bueno de CUALQUIER ticker individual, sin importar en qué
    // categoría/fila esté hoy.
    const preciosViejosPorTicker = {};

    if (ultimaFilaActual >= FILA_INICIO) {
        const rango = hojaPrecios.getRange(FILA_INICIO, 1, ultimaFilaActual - FILA_INICIO + 1, 2);
        const dataVieja = rango.getValues();
        const backgroundViejo = rango.getBackgrounds();
        const numberFormatsViejos = rango.getNumberFormats();

        let categoriaActual = null;
        for (let i = 0; i < dataVieja.length; i++) {
            const val = String(dataVieja[i][0] || "").trim();
            if (ordenCategorias.map(c => c.toUpperCase()).includes(val.toUpperCase())) {
                categoriaActual = ordenCategorias.find(c => c.toUpperCase() === val.toUpperCase());
            }
            if (categoriaActual) {
                bloquesViejosPorCategoria[categoriaActual].push({
                    val: dataVieja[i],
                    bg: backgroundViejo[i],
                    fmt: numberFormatsViejos[i]
                });

                // Si esta fila tiene un ticker real (no un header/subheader) con
                // precio numérico > 0, lo guardamos como "último precio bueno"
                const posibleTicker = String(dataVieja[i][0] || "").trim().toUpperCase();
                const posiblePrecio = dataVieja[i][1];
                const esHeaderOSubheader = (posibleTicker === categoriaActual.toUpperCase()) || posibleTicker === "TICKER";
                if (!esHeaderOSubheader && posibleTicker && typeof posiblePrecio === 'number' && posiblePrecio > 0) {
                    preciosViejosPorTicker[posibleTicker] = {
                        precio: posiblePrecio,
                        fmt: numberFormatsViejos[i][1]
                    };
                }
            }
        }
    }

    if (ultimaFilaActual >= FILA_INICIO) {
        hojaPrecios.getRange(FILA_INICIO, 1, ultimaFilaActual - FILA_INICIO + 1, 2).clearContent().clearFormat();
    }

    const AZUL_HEADER = "#1a73e8";
    const BLANCO = "#ffffff";
    const AMARILLO_CLARO = "#fff9e6";
    const AMARILLO_ALT = "#fffdf5";
    const GRIS_SUBHEADER = "#f0f0f0";

    let filaActual = FILA_INICIO;
    let tickersConservados = [];

    ordenCategorias.forEach((tipo, idx) => {
        if (fetchOk[tipo]) {
            const tickers = [...grupos[tipo]].sort();
            if (tickers.length === 0) return;

            const rHeader = hojaPrecios.getRange(filaActual, 1, 1, 2);
            rHeader.merge().setValue(tipo.toUpperCase())
                .setBackground(AZUL_HEADER).setFontColor(BLANCO).setFontWeight("bold")
                .setFontSize(10).setFontFamily("Arial").setHorizontalAlignment("center");
            filaActual++;

            hojaPrecios.getRange(filaActual, 1).setValue("Ticker");
            hojaPrecios.getRange(filaActual, 2).setValue("Último Operado");
            hojaPrecios.getRange(filaActual, 1, 1, 2)
                .setBackground(GRIS_SUBHEADER).setFontWeight("bold").setFontFamily("Arial")
                .setFontSize(10).setHorizontalAlignment("center");
            filaActual++;

            tickers.forEach((ticker, i) => {
                let precio = precios[ticker] !== undefined ? precios[ticker] : "";

                // FIX: si el precio nuevo es 0 o no vino, y tenemos un precio
                // bueno anterior guardado para este ticker puntual, lo usamos
                // en vez de pisarlo con un cero (mercado cerrado / pre-market).
                if ((precio === 0 || precio === "") && preciosViejosPorTicker[ticker]) {
                    precio = preciosViejosPorTicker[ticker].precio;
                    tickersConservados.push(ticker);
                }

                const fondo = i % 2 === 0 ? AMARILLO_CLARO : AMARILLO_ALT;

                hojaPrecios.getRange(filaActual, 1).setValue(ticker)
                    .setBackground(fondo).setFontFamily("Arial").setFontSize(10)
                    .setFontWeight("normal").setHorizontalAlignment("left");
                hojaPrecios.getRange(filaActual, 2).setValue(precio)
                    .setBackground(fondo).setFontFamily("Arial").setFontSize(10)
                    .setNumberFormat("#,##0.00").setHorizontalAlignment("right");
                filaActual++;
            });
        } else {
            const bloqueViejo = bloquesViejosPorCategoria[tipo];
            if (bloqueViejo && bloqueViejo.length > 0) {
                bloqueViejo.forEach(item => {
                    hojaPrecios.getRange(filaActual, 1).setValue(item.val[0]).setBackground(item.bg[0]);
                    hojaPrecios.getRange(filaActual, 2).setValue(item.val[1]).setBackground(item.bg[1]);
                    if (item.fmt[1]) hojaPrecios.getRange(filaActual, 2).setNumberFormat(item.fmt[1]);
                    filaActual++;
                });
            }
        }

        if (idx < ordenCategorias.length - 1) filaActual++;
    });

    hojaPrecios.setColumnWidth(1, 120);
    hojaPrecios.setColumnWidth(2, 140);

    if (tickersConservados.length > 0) {
        console.log(`ℹ️ Precios conservados (dato nuevo era 0/ausente): ${tickersConservados.join(', ')}`);
    }
    console.log(`✅ Escritura finalizada | Categorías OK: ${Object.keys(fetchOk).filter(k => fetchOk[k]).join(', ')} | ${Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "dd/MM/yyyy HH:mm:ss")}`);
}

// =================================================================================
// 4. INSTALADORES DE TRIGGERS
// =================================================================================
function instalarTriggerCCL() {
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'REGISTRAR_CCL_INTELIGENTE').forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('REGISTRAR_CCL_INTELIGENTE').timeBased().everyMinutes(30).create();
    console.log("✅ Trigger CCL instalado: cada 30 minutos.");
}

// FIX: reemplaza a instalarTriggerCAFCI() e instalarTriggerCompass() — las dos
// versiones viejas apuntaban al nombre de función incorrecto/viejo
// ('ACTUALIZAR_PRECIOS_COMPASS', que ya no existe — la función real se llama
// ACTUALIZAR_PRECIOS_CAFCI). Además, .onWeekDay() encadenado NO se acumula:
// cada llamada pisa a la anterior, así que terminaba corriendo un solo día
// (el último de la cadena). Acá se crean 5 triggers independientes, uno por
// día hábil, todos apuntando al nombre correcto.
function instalarTriggerPreciosCAFCI() {
    ScriptApp.getProjectTriggers()
        .filter(t => t.getHandlerFunction() === 'ACTUALIZAR_PRECIOS_CAFCI'
                   || t.getHandlerFunction() === 'ACTUALIZAR_PRECIOS_COMPASS')
        .forEach(t => ScriptApp.deleteTrigger(t));

    const dias = [
        ScriptApp.WeekDay.MONDAY,
        ScriptApp.WeekDay.TUESDAY,
        ScriptApp.WeekDay.WEDNESDAY,
        ScriptApp.WeekDay.THURSDAY,
        ScriptApp.WeekDay.FRIDAY
    ];

    dias.forEach(dia => {
        ScriptApp.newTrigger('ACTUALIZAR_PRECIOS_CAFCI')
            .timeBased()
            .onWeekDay(dia)
            .atHour(20)
            .create();
    });

    console.log("✅ Trigger Precios CAFCI instalado: 5 triggers independientes, lunes a viernes 20:30hs, apuntando a ACTUALIZAR_PRECIOS_CAFCI.");
}

function instalarTriggerData912() {
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actualizarPrecios').forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('actualizarPrecios').timeBased().everyMinutes(30).create();
    console.log("✅ Trigger data912 instalado: cada 30 minutos.");
}