// =================================================================================
// ===  VALIDACIONMOTORNUEVO.GS — Script standalone, de SOLO LECTURA            ===
// ===  Paso 2 de la consolidación. Compara, ticker por ticker, la posición     ===
// ===  (q y costo) que calcula el sistema VIEJO (calcularCantidadesNetas, ya   ===
// ===  en Motor.gs) contra la que calcula el motor NUEVO (snapshotAFecha, en   ===
// ===  MotorPosiciones.gs) para HOY.                                          ===
// ===                                                                          ===
// ===  No modifica el Log ni ningún consumidor real. Único efecto secundario:  ===
// ===  crea la hoja temporal "VALIDACION_MOTOR_TEMP". Se puede borrar junto    ===
// ===  con este archivo una vez confirmado el resultado.                      ===
// =================================================================================

function VALIDAR_MOTOR_NUEVO_VS_VIEJO() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const logSheet = ss.getSheetByName(HOJAS.LOG);
    if (!logSheet) {
        ui.alert('No se encontró la hoja ' + HOJAS.LOG + '.');
        return;
    }

    const logData = logSheet.getDataRange().getValues();
    const logSinHeader = logData.slice(1);
    logSinHeader.sort((a, b) => new Date(a[1]) - new Date(b[1]));

    // --- Sistema VIEJO: calcularCantidadesNetas(), ya vive en Motor.gs ---
    const viejo = calcularCantidadesNetas();

    // --- Sistema NUEVO: snapshotAFecha(), en MotorPosiciones.gs ---
    const HOY = new Date();
    HOY.setHours(23, 59, 59, 999); // corte inclusive de todo el día de hoy
    const nuevo = snapshotAFecha(logSinHeader, HOY);

    // --- Comparación ---
    const TOLERANCIA = 0.01; // un centavo de diferencia se considera "igual"
    const todosLosTickers = new Set([
        ...Object.keys(viejo),
        ...Object.keys(nuevo.posiciones)
    ]);

    let filas = [];
    let huboDiferencias = false;

    todosLosTickers.forEach(tk => {
        const v = viejo[tk] || { q: 0, costo: 0 };
        const n = nuevo.posiciones[tk] || { q: 0, costo: 0 };

        const diffQ = Math.abs((v.q || 0) - (n.q || 0));
        const diffCosto = Math.abs((v.costo || 0) - (n.costo || 0));
        const coincide = diffQ <= TOLERANCIA && diffCosto <= TOLERANCIA;

        if (!coincide) huboDiferencias = true;

        // Solo mostramos tickers con posición viva en alguno de los dos sistemas,
        // o con diferencia (para no llenar la hoja de ceros irrelevantes).
        if ((v.q || 0) > 0.0001 || (n.q || 0) > 0.0001 || !coincide) {
            filas.push([
                tk,
                v.q || 0, n.q || 0, diffQ,
                v.costo || 0, n.costo || 0, diffCosto,
                coincide ? 'OK' : '⚠️ DIFIERE'
            ]);
        }
    });

    // Ordenar: las diferencias primero, para que salten a la vista.
    filas.sort((a, b) => {
        if (a[7] === b[7]) return String(a[0]).localeCompare(String(b[0]));
        return a[7] === '⚠️ DIFIERE' ? -1 : 1;
    });

    let hojaSalida = ss.getSheetByName('VALIDACION_MOTOR_TEMP');
    if (hojaSalida) ss.deleteSheet(hojaSalida);
    hojaSalida = ss.insertSheet('VALIDACION_MOTOR_TEMP');

    hojaSalida.getRange(1, 1, 1, 8).setValues([[
        'Ticker', 'Q (viejo)', 'Q (nuevo)', 'Dif. Q',
        'Costo (viejo)', 'Costo (nuevo)', 'Dif. Costo', 'Resultado'
    ]]).setFontWeight('bold');

    if (filas.length > 0) {
        hojaSalida.getRange(2, 1, filas.length, 8).setValues(filas);
        hojaSalida.getRange(2, 5, filas.length, 3).setNumberFormat('$#,##0.00');
    }

    // Resaltar filas que difieren.
    const rangoResultado = hojaSalida.getRange(2, 8, Math.max(filas.length, 1), 1);
    const reglaDiff = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('DIFIERE')
        .setBackground('#ffebee').setFontColor('#c62828').setBold(true)
        .setRanges([rangoResultado]).build();
    hojaSalida.setConditionalFormatRules([reglaDiff]);
    hojaSalida.autoResizeColumns(1, 8);

    // --- Advertencias que generó el motor nuevo (tope de venta, sin stock, etc.) ---
    let mensajeAdvertencias = '';
    if (nuevo.advertencias.length > 0) {
        mensajeAdvertencias = '\n\n⚠️ El motor nuevo emitió ' + nuevo.advertencias.length +
            ' advertencia(s) de reglas defensivas (esperable si ya sabés de casos como BRKB/HMY). ' +
            'Quedan detalladas en la hoja "VALIDACION_ADVERTENCIAS_TEMP".';

        let hojaAdv = ss.getSheetByName('VALIDACION_ADVERTENCIAS_TEMP');
        if (hojaAdv) ss.deleteSheet(hojaAdv);
        hojaAdv = ss.insertSheet('VALIDACION_ADVERTENCIAS_TEMP');
        hojaAdv.getRange(1, 1, 1, 3).setValues([['Ticker', 'Fecha', 'Mensaje']]).setFontWeight('bold');
        const filasAdv = nuevo.advertencias.map(a => [a.ticker, a.fecha, a.mensaje]);
        hojaAdv.getRange(2, 1, filasAdv.length, 3).setValues(filasAdv);
        hojaAdv.getRange(2, 2, filasAdv.length, 1).setNumberFormat('dd/MM/yyyy');
        hojaAdv.autoResizeColumns(1, 3);
    }

    const mensajeFinal = huboDiferencias
        ? '⚠️ Hay diferencias entre el motor viejo y el nuevo para ' +
          filas.filter(f => f[7] === '⚠️ DIFIERE').length +
          ' ticker(s). Revisá la hoja "VALIDACION_MOTOR_TEMP" antes de seguir.' +
          mensajeAdvertencias
        : '✅ El motor nuevo coincide con el viejo en q y costo para todos los tickers (tolerancia $' +
          TOLERANCIA + ').' + mensajeAdvertencias;

    ui.alert(mensajeFinal);
    console.log(mensajeFinal);
}
