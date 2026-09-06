// =================================================================================
// ===  TITANIUM v2 — MOTORPOSICIONES.GS                                        ===
// ===  Motor único de posiciones. Fuente de verdad de "cuánto tengo de cada    ===
// ===  ticker y a qué costo", sea a hoy o a cualquier fecha de corte pasada.   ===
// ===                                                                          ===
// ===  PASO 1 de la consolidación (ver debate en chat "Portafolio Ultimate").  ===
// ===  Este archivo es NUEVO y AUTÓNOMO: no reemplaza ni borra ningún loop     ===
// ===  existente en Motor.gs / Auditoria.gs / Fiscal.gs. El sistema sigue      ===
// ===  funcionando exactamente igual que antes hasta que migremos cada         ===
// ===  consumidor uno por uno (Paso 3), y solo después de validar (Paso 2)     ===
// ===  que este motor da los mismos números que el sistema viejo.              ===
// ===                                                                          ===
// ===  Reutiliza cleanNum() y normalizar(), ya definidas en Motor.gs.          ===
// =================================================================================


// ---------------------------------------------------------------------------------
// 1. LA ÚNICA PUERTA DE ENTRADA PARA LAS REGLAS DE NEGOCIO
// ---------------------------------------------------------------------------------
// aplicarMovimiento() es la ÚNICA función de todo el sistema que sabe qué le hace
// una fila del Log a una posición. Cualquier regla de negocio sobre cómo se
// actualiza q/costo vive ACÁ, y solo acá. Si el día de mañana hay que ajustar
// una regla, se toca en un solo lugar y corrige automáticamente a todos los
// consumidores (Dashboard, Auditoría, Fiscal, Bombonera, trigger de 15').
//
// Reglas de negocio ya validadas/decididas y CERRADAS (no reabrir sin motivo):
//
//   a) AMORTIZACIÓN — Método B (ya validado en producción, no se reabre):
//        - Nunca toca `q` (los nominales no bajan por una amortización parcial).
//        - `costo` se reduce en min(costo, montoUSD) — nunca puede quedar negativo.
//        - `costoOriginal` queda INTACTO siempre (es el costo histórico de entrada,
//          no se toca nunca por una amortización).
//        - Si `montoUSD` viene en 0 (fila mal cargada o placeholder), la fila se
//          SALTEA por completo (no se resta nada, no se suma a cobrado) y se
//          genera un warning para revisión manual — nunca se aplica en silencio.
//
//   b) TOPE DE VENTA A CANTIDAD TENIDA (nueva regla defensiva, permanente):
//        Si una venta pide más cantidad de la que hay en la posición, se topea
//        la venta a la cantidad realmente disponible y se genera un warning.
//        Encontrado en el Log real (BRKB, HMY) — ver diagnóstico del 06/09/2026.
//
//   c) VENTA SIN STOCK PREVIO = GANANCIA REALIZADA TOTAL (nueva regla defensiva,
//      permanente):
//        Si se vende/rescata un ticker con saldo 0 (posición heredada, fila de
//        compra faltante, etc.), el 100% del monto se toma como ganancia
//        realizada, ya que no hay costo contra el cual restarlo. Se genera un
//        warning.
//
// Decisión de diseño (confirmada en debate del 06/09/2026):
//   El campo `tipo` de la posición se actualiza con el ÚLTIMO Tipo_Activo NO
//   VACÍO visto, no con el de la primera compra ni con cualquier valor a
//   ciegas. Concretamente:
//     - Si la fila trae Tipo_Activo con contenido, se usa y pisa el anterior.
//     - Si la fila viene con Tipo_Activo vacío, se conserva el tipo que ya
//       tenía la posición (no se pisa por una fila incompleta).
//   Esto permite corregir el tipo de un ticker cargando bien una fila nueva,
//   sin perder el tipo por una carga a mano incompleta. Reemplaza tanto el
//   comportamiento de calcularCantidadesNetas() (pisaba con cualquier valor,
//   vacío incluido) como el de generarDatosMaestros() (solo tomaba la primera
//   fila) — ambos eran divergencias silenciosas, ya unificadas acá.
//
// @param {Object} estado  Posición actual del ticker: { q, costo, costoOriginal,
//                          cobrado, wDate, tipo }. Para un ticker nuevo, usar
//                          crearEstadoInicial().
// @param {Array}  fila    Una fila cruda del Log (mismo array que devuelve
//                          getDataRange().getValues(), con los mismos índices
//                          que ya usa el resto del sistema).
// @return {Object} { estado: <posición actualizada>, detalle: <qué pasó en esta fila> }
//
function aplicarMovimiento(estado, fila) {
    const fechaFila = fila[1] instanceof Date ? fila[1] : new Date(fila[1]);
    const ticker = String(fila[3]).toUpperCase().trim();
    const tipoStr = String(fila[4]);
    const mov = String(fila[5]).toLowerCase().trim();
    const cant = cleanNum(fila[6]);
    const montoUSD = Math.abs(cleanNum(fila[10]));
    const ratioSplit = cleanNum(fila[11]);

    // Copia superficial: no mutamos el objeto que nos pasaron.
    let e = {
        q: estado.q, costo: estado.costo, costoOriginal: estado.costoOriginal,
        cobrado: estado.cobrado, wDate: estado.wDate, tipo: estado.tipo
    };

    let detalle = {
        fecha: fechaFila, ticker: ticker, tipo: tipoStr, movimiento: mov,
        montoUSD: montoUSD, cantidad: cant,
        esCompra: false, esVenta: false, esRenta: false, esDividendo: false,
        esAmortizacion: false, esSplit: false,
        gananciaRealizada: 0, gananciaExtra: 0, cashflow: 0,
        warning: null, skip: false
    };

    // Ver "Decisión de diseño" arriba: último Tipo_Activo NO VACÍO gana; una
    // fila con Tipo_Activo vacío nunca borra el tipo que ya tenía la posición.
    if (tipoStr && tipoStr.trim() !== '') e.tipo = tipoStr;

    if (mov.includes('compra') || mov.includes('aporte') || mov.includes('suscripcion') || mov.includes('canje_entrada')) {
        detalle.esCompra = true;

        const fechaMs = fechaFila.getTime();
        if (e.q <= 0 || !e.wDate) e.wDate = fechaMs;
        else e.wDate = ((e.q * e.wDate) + (cant * fechaMs)) / (e.q + cant);

        e.q += cant;
        e.costo += montoUSD;
        e.costoOriginal += montoUSD;
        detalle.cashflow = -montoUSD;

    } else if (mov.includes('venta') || mov.includes('rescate') || mov.includes('canje_salida')) {
        detalle.esVenta = true;

        if (e.q <= 0) {
            // Regla (c): venta sin stock previo = ganancia realizada total.
            detalle.gananciaRealizada = montoUSD;
            detalle.cashflow = montoUSD;
            detalle.warning = 'Venta sin stock previo registrado (' + ticker + ', fecha ' +
                Utilities.formatDate(fechaFila, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
                ') — se tomó el 100% del monto como ganancia realizada.';
        } else {
            let cantVenta = cant;
            if (cant > e.q) {
                // Regla (b): tope de venta a la cantidad realmente tenida.
                detalle.warning = 'Venta de ' + ticker + ' (fecha ' +
                    Utilities.formatDate(fechaFila, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
                    ') supera el saldo tenido: se pidieron ' + cant + ' contra un saldo de ' +
                    e.q + '. Se topeó la venta a ' + e.q + '.';
                cantVenta = e.q;
            }

            const ppc = e.costo / e.q;
            const ppcOriginal = e.costoOriginal / e.q;
            const costoVenta = ppc * cantVenta;
            const costoOriginalVenta = ppcOriginal * cantVenta;
            const ganancia = montoUSD - costoVenta;

            detalle.gananciaRealizada = ganancia;
            detalle.cashflow = montoUSD;

            e.q -= cantVenta;
            e.costo -= costoVenta;
            e.costoOriginal -= costoOriginalVenta;
            if (e.q < 0.0001) { e.q = 0; e.costo = 0; e.costoOriginal = 0; }
        }

    } else if (mov.includes('dividendo') || mov.includes('renta') || mov.includes('interes')) {
        detalle.esRenta = true;
        detalle.esDividendo = mov.includes('dividendo');

        e.cobrado += montoUSD;
        detalle.cashflow = montoUSD;
        detalle.gananciaRealizada = montoUSD; // toda renta cobrada es ganancia realizada

    } else if (mov.includes('amortiza')) {
        detalle.esAmortizacion = true;

        if (montoUSD <= 0) {
            // Regla (a), último punto: Monto=0 -> se saltea la fila y se avisa.
            detalle.skip = true;
            detalle.warning = 'Amortización de ' + ticker + ' (fecha ' +
                Utilities.formatDate(fechaFila, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
                ') con Monto_Neto_USD = 0 — fila ignorada, revisar la carga.';
            return { estado: e, detalle: detalle };
        }

        e.cobrado += montoUSD;
        const reduccionCosto = Math.min(montoUSD, e.costo);
        const gananciaExtra = Math.max(0, montoUSD - e.costo);

        e.costo -= reduccionCosto;               // Método B: nunca negativo.
        // e.costoOriginal: intacto a propósito, no se toca acá. Método B.

        detalle.gananciaExtra = gananciaExtra;
        detalle.gananciaRealizada = gananciaExtra; // lo cobrado por sobre el costo vivo, ya realizado
        detalle.cashflow = montoUSD;

    } else if (mov.includes('split')) {
        detalle.esSplit = true;
        if (ratioSplit > 0 && e.q > 0) {
            e.q = e.q * ratioSplit;
        }
    }

    return { estado: e, detalle: detalle };
}

// Estado inicial en blanco para un ticker que aparece por primera vez.
function crearEstadoInicial(tipoStr) {
    return { q: 0, costo: 0, costoOriginal: 0, cobrado: 0, wDate: null, tipo: tipoStr || '' };
}


// ---------------------------------------------------------------------------------
// 2. "¿CUÁNTO TENÍA DE CADA COSA A TAL FECHA?"
// ---------------------------------------------------------------------------------
// Recorre el Log aplicando aplicarMovimiento() a cada fila hasta (e incluyendo)
// fechaCorte, y devuelve la posición de cada ticker en ese momento.
//
// Reemplaza en el futuro (Paso 3, uno por uno, con validación previa):
//   - calcularCantidadesNetas()          -> snapshotAFecha(log, HOY)
//   - el mini-loop de calcularRendimientoLimpioYTD() (tenenciasCorte) -> snapshotAFecha(log, fechaDeCorteQueYaUsa)
//   - el inventario de GENERAR_INFORME_BIENES_PERSONALES() -> snapshotAFecha(log, 31/12 23:59:59 del año fiscal elegido)
//
// USD/CASH queda EXCLUIDO de las posiciones (igual que en todo el sistema hoy) —
// la caja se calcula aparte con calcularCajaVirtual(). Fiscal.gs, que necesita
// el monto de USD/CASH para su propio reporte, lo sigue resolviendo con su
// propia lógica de caja, por fuera de este motor.
//
// @param {Array} logSinHeader  Filas del Log SIN el header, YA ordenadas por
//                                fecha ascendente (igual que hace hoy cada
//                                consumidor antes de procesar).
// @param {Date}  fechaCorte    Fecha límite inclusive. Para "hoy" pasar HOY_SIMULADA.
//                                Para Fiscal, pasar new Date(anioFiscal, 11, 31, 23, 59, 59).
// @return {Object} { posiciones: { TICKER: {q, costo, costoOriginal, cobrado, wDate, tipo} },
//                     advertencias: [ { ticker, fecha, mensaje } ] }
//
function snapshotAFecha(logSinHeader, fechaCorte) {
    const fechaCorteMs = fechaCorte.getTime();
    let posiciones = {};
    let advertencias = [];

    (logSinHeader || []).forEach(fila => {
        const fechaFila = fila[1] instanceof Date ? fila[1] : new Date(fila[1]);
        if (isNaN(fechaFila.getTime()) || fechaFila.getTime() > fechaCorteMs) return;

        const ticker = String(fila[3]).toUpperCase().trim();
        if (!ticker || ticker === 'USD' || ticker === 'CASH') return;

        if (!posiciones[ticker]) posiciones[ticker] = crearEstadoInicial(String(fila[4]));

        const resultado = aplicarMovimiento(posiciones[ticker], fila);
        posiciones[ticker] = resultado.estado;

        if (resultado.detalle.warning) {
            advertencias.push({
                ticker: ticker, fecha: fechaFila, mensaje: resultado.detalle.warning
            });
        }
    });

    return { posiciones: posiciones, advertencias: advertencias };
}


// ---------------------------------------------------------------------------------
// 3. "¿QUÉ PASÓ, TRANSACCIÓN POR TRANSACCIÓN, Y CUÁNTO GANÉ REALIZADO POR AÑO?"
// ---------------------------------------------------------------------------------
// Recorre TODO el Log aplicando aplicarMovimiento(), y además de la posición
// final guarda el detalle de cada transacción y los totales agrupados por año.
//
// Reemplaza en el futuro (Paso 3):
//   - las estadísticas de dividendos/renta de generarDatosMaestros()
//   - el loop completo de GENERAR_AUDITORIA_COMPLETA() (que hoy arma su propio
//     transDetalle y gciaPorAnio)
//
// @param {Array} logSinHeader  Filas del Log SIN header, ya ordenadas por fecha.
// @return {Object} {
//   posiciones:    posición final de cada ticker a la fecha del último movimiento,
//   transacciones: [ { fecha, ticker, tipo, movimiento, montoUSD, gananciaRealizada, anio } ],
//   statsPorAnio:  { anio: { gciaCapRV, gciaCapRF, divRV, rentaRF, amortizRF } },
//   advertencias:  [ { ticker, fecha, mensaje } ]
// }
//
function ledgerCompleto(logSinHeader) {
    let posiciones = {};
    let transacciones = [];
    let statsPorAnio = {};
    let advertencias = [];

    (logSinHeader || []).forEach(fila => {
        const fechaFila = fila[1] instanceof Date ? fila[1] : new Date(fila[1]);
        if (isNaN(fechaFila.getTime())) return;

        const ticker = String(fila[3]).toUpperCase().trim();
        if (!ticker || ticker === 'USD' || ticker === 'CASH') return;

        if (!posiciones[ticker]) posiciones[ticker] = crearEstadoInicial(String(fila[4]));

        const resultado = aplicarMovimiento(posiciones[ticker], fila);
        posiciones[ticker] = resultado.estado;
        const d = resultado.detalle;

        if (d.warning) {
            advertencias.push({ ticker: ticker, fecha: fechaFila, mensaje: d.warning });
        }
        if (d.skip) return; // p.ej. amortización con Monto=0

        const anio = fechaFila.getFullYear();
        const esRV = d.tipo.toLowerCase().includes('cedear') || d.tipo.toLowerCase().includes('accion');

        transacciones.push({
            fecha: fechaFila, ticker: ticker, tipo: d.tipo, movimiento: d.movimiento,
            montoUSD: d.montoUSD, gananciaRealizada: d.gananciaRealizada, anio: anio
        });

        if (!statsPorAnio[anio]) {
            statsPorAnio[anio] = { gciaCapRV: 0, gciaCapRF: 0, divRV: 0, rentaRF: 0, amortizRF: 0 };
        }
        const s = statsPorAnio[anio];

        if (d.esVenta) {
            if (esRV) s.gciaCapRV += d.gananciaRealizada;
            else s.gciaCapRF += d.gananciaRealizada;
        } else if (d.esRenta) {
            if (d.esDividendo) s.divRV += d.montoUSD;
            else s.rentaRF += d.montoUSD;
        } else if (d.esAmortizacion) {
            s.amortizRF += d.montoUSD;
            if (d.gananciaExtra > 0) s.rentaRF += d.gananciaExtra;
        }
    });

    return {
        posiciones: posiciones, transacciones: transacciones,
        statsPorAnio: statsPorAnio, advertencias: advertencias
    };
}