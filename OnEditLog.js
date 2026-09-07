// =================================================================================
// ===  ONEDITLOG.GS — Autogenerar ID_Movimiento y extender la fila siguiente  ===
// =================================================================================
// Feature acordada al arranque del proyecto "Portafolio Ultimate" (ver debate).
//
// QUÉ HACE:
//   1. Cuando cargás el Ticker (columna D) de una fila nueva en el Log, si esa
//      fila todavía no tiene ID_Movimiento (columna A), le genera uno — un
//      UUID, mismo formato que ya usás — y lo escribe como TEXTO FIJO.
//   2. Extiende el formato visual y las validaciones de datos (dropdowns) de
//      esa fila hacia la fila de abajo, para que ya te quede lista para
//      cargar el próximo movimiento sin copiar/pegar nada a mano.
//
// POR QUÉ EL ID ES FIJO Y NUNCA SE RECALCULA (decisión ya tomada, no reabrir):
//   La columna ID_Relacionado está pensada para vincular movimientos entre sí
//   en casos de split, contra-split o canje de un bono por otro. Si el ID de
//   la fila referenciada pudiera cambiar de valor con el tiempo (como pasaría
//   con una fórmula), esos vínculos se romperían. Por eso va por script, con
//   setValue() de un string fijo — nunca una fórmula, nunca se recalcula, y
//   nunca se pisa si la fila ya tenía uno.
//
// Requiere trigger SIMPLE — no hace falta instalar nada manualmente, Apps
// Script lo dispara solo con que exista esta función.
// =================================================================================

function onEdit(e) {
    try {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== HOJAS.LOG) return;

        const COL_ID = 1;      // A - ID_Movimiento
        const COL_TICKER = 4;  // D - Ticker

        const filaInicio = e.range.getRow();
        const filaFin = e.range.getLastRow();
        const colInicio = e.range.getColumn();
        const colFin = e.range.getLastColumn();

        // Solo nos interesa si la edición tocó la columna Ticker (cubre tanto
        // tipear una celda como pegar un bloque de varias filas/columnas).
        if (colInicio > COL_TICKER || colFin < COL_TICKER) return;

        for (let fila = Math.max(filaInicio, 2); fila <= filaFin; fila++) {
            const celdaTicker = sheet.getRange(fila, COL_TICKER);
            const ticker = String(celdaTicker.getValue()).trim();
            if (!ticker) continue; // no generamos ID si el ticker está vacío

            const celdaId = sheet.getRange(fila, COL_ID);
            if (String(celdaId.getValue()).trim() === '') {
                celdaId.setValue(Utilities.getUuid());
            }

            extenderFilaSiguienteLog(sheet, fila);
        }
    } catch (err) {
        console.error('Error en onEdit del Log: ' + err.message);
    }
}

function extenderFilaSiguienteLog(sheet, filaOrigen) {
    const COL_TICKER = 4;
    const filaDestino = filaOrigen + 1;
    const ultimaColumna = sheet.getLastColumn();

    // Si la fila siguiente YA tiene ticker cargado, no la tocamos — evita
    // pisar una fila que ya está en uso (por ejemplo, si el usuario está
    // editando una fila del medio del historial, no la de más abajo del todo).
    const tickerDestino = String(sheet.getRange(filaDestino, COL_TICKER).getValue()).trim();
    if (tickerDestino !== '') return;

    const rangoOrigen = sheet.getRange(filaOrigen, 1, 1, ultimaColumna);
    const rangoDestino = sheet.getRange(filaDestino, 1, 1, ultimaColumna);

    // PASTE_FORMAT: colores, bordes, formato de fecha/número — todo el
    // aspecto visual de la fila, sin tocar ningún valor.
    rangoOrigen.copyTo(rangoDestino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    // PASTE_DATA_VALIDATION: copia las reglas de validación (los dropdowns)
    // de TODA la fila — no hace falta hardcodear qué columnas tienen
    // dropdown hoy (Broker, Tipo_Activo, Tipo_Movimiento, Moneda_ORIG); si el
    // día de mañana agregás una validación en otra columna, se extiende sola,
    // sin tener que tocar este script.
    rangoOrigen.copyTo(rangoDestino, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}