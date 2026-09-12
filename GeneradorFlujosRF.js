// =================================================================================
// ===  GENERADORFLUJOSRF.GS — Creador Automático de Hojas de Bonos y ONs        ===
// =================================================================================
// Genera hojas de flujo de fondos con:
//  - Encabezado fila 1 color #356854, texto blanco negrita.
//  - Resaltado de fila de próximo pago en Magenta Claro 3 (#ead1dc).
//  - Metadatos en K1 (Ticker USD), L1 (Ticker ARS), M1 (Frecuencia anual).
//  - Fórmula canónica en K3: =IFERROR(VLOOKUP(L1, Cartera!A:C, 3, FALSE), 100)
//    (Google Sheets la traduce en español con punto y coma: =SI.ERROR(BUSCARV(L1; Cartera!A:C; 3; FALSO); 100)).
//  - Fórmulas vivas en G (% Residual), H (# Renta), I (# Capital), J (# Total).
//  - Registro en Precios!S:T ÚNICAMENTE si hay compra real (tenencia > 0 y != 100).
// =================================================================================

function abrirDialogoNuevoBono() {
    const html = HtmlService.createHtmlOutputFromFile('DialogoNuevoBono')
        .setWidth(980)
        .setHeight(750)
        .setTitle('TITANIUM — Generador de Flujos de Renta Fija (Bono / ON)');
    SpreadsheetApp.getUi().showModalDialog(html, 'TITANIUM — Generador de Flujos (Bono / ON)');
}

function parseDateYMD(str) {
    if (!str) return new Date();
    const parts = String(str).split('T')[0].split('-');
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 0, 0, 0);
}

function crearHojaBonoBackend(config) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const tickerUSD = String(config.tickerUSD || '').toUpperCase().trim();
        const tickerPesos = String(config.tickerPesos || '').toUpperCase().trim();
        const frecuencia = parseInt(config.frecuencia, 10) || 2;
        const cupones = config.cupones || [];

        if (!tickerUSD) throw new Error('El Ticker en Dólares (USD) es obligatorio.');
        if (cupones.length === 0) throw new Error('No se generaron cupones para el bono.');

        let sheet = ss.getSheetByName(tickerUSD);
        if (sheet) {
            throw new Error('Ya existe una hoja con el nombre "' + tickerUSD + '". Eliminala o cámbiale el nombre para volver a generarla.');
        }

        sheet = ss.insertSheet(tickerUSD);

        // 1. Encabezados de Columnas A a J
        const headers = [
            '# Cupon', 'Vencimiento', 'Fecha de pago', '% Interés',
            '% Capitalización', '% Amortización', '% Residual',
            '# Renta', '# Capital', '# Total'
        ];

        sheet.getRange(1, 1, 1, 10).setValues([headers])
            .setBackground('#356854')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontSize(10)
            .setFontFamily('Arial')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');

        sheet.setRowHeight(1, 28);

        // 2. Metadatos K1, L1, M1
        sheet.getRange('K1').setValue(tickerUSD).setFontWeight('bold').setHorizontalAlignment('center');
        sheet.getRange('L1').setValue(tickerPesos).setFontWeight('bold').setHorizontalAlignment('center');
        sheet.getRange('M1').setValue(frecuencia).setFontWeight('bold').setHorizontalAlignment('center');

        // 3. Tenencia Viva en K3
        sheet.getRange('K3').setFormula('=IFERROR(VLOOKUP(L1, Cartera!A:C, 3, FALSE), 100)')
            .setNumberFormat('#,##0')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');

        // 4. Cupón 0 (Fila 2)
        const fechaEmision = parseDateYMD(config.fechaEmision);
        sheet.getRange('A2').setValue(0).setHorizontalAlignment('center');
        sheet.getRange('C2').setValue(fechaEmision).setNumberFormat('d/M/yyyy').setHorizontalAlignment('center');
        sheet.getRange('G2').setValue(1).setNumberFormat('0.00%').setHorizontalAlignment('center');
        sheet.getRange('J2').setValue(0).setNumberFormat('#,##0.00').setHorizontalAlignment('right');

        // 5. Filas de Cupones (Fila 3 en adelante)
        const hoy = new Date();
        hoy.setHours(0, 0, 0, 0);
        let proximaFilaPago = null;

        const numFilas = cupones.length;
        const valoresA_F = [];
        const formulasG_J = [];

        for (let i = 0; i < numFilas; i++) {
            const c = cupones[i];
            const filaActual = i + 3;
            const filaAnterior = filaActual - 1;

            const vto = parseDateYMD(c.vencimiento);
            const pago = parseDateYMD(c.fechaPago);

            if (!proximaFilaPago && pago.getTime() >= hoy.getTime()) {
                proximaFilaPago = filaActual;
            }

            valoresA_F.push([
                i + 1,
                vto,
                pago,
                (Number(c.tasaInteres) || 0) / 100,
                (Number(c.capitalizacion) || 0) / 100,
                (Number(c.amortizacion) || 0) / 100
            ]);

            formulasG_J.push([
                '=G' + filaAnterior + '-F' + filaActual + '+E' + filaActual,
                '=$K$3*D' + filaActual + '/$M$1*G' + filaAnterior,
                '=$K$3*F' + filaActual,
                '=H' + filaActual + '+I' + filaActual
            ]);
        }

        // Escribir datos fijos y fórmulas
        sheet.getRange(3, 1, numFilas, 6).setValues(valoresA_F);
        sheet.getRange(3, 7, numFilas, 4).setFormulas(formulasG_J);

        // 6. Formatos Numéricos y Alineaciones
        sheet.getRange(3, 1, numFilas, 1).setNumberFormat('0').setHorizontalAlignment('center');
        sheet.getRange(3, 2, numFilas, 2).setNumberFormat('d/M/yyyy').setHorizontalAlignment('center');
        sheet.getRange(3, 4, numFilas, 4).setNumberFormat('0.00%').setHorizontalAlignment('center');
        sheet.getRange(3, 8, numFilas, 3).setNumberFormat('#,##0.00').setHorizontalAlignment('right');

        // Bordes suaves
        sheet.getRange(1, 1, numFilas + 2, 10).setBorder(true, true, true, true, true, true, '#d0d7de', SpreadsheetApp.BorderStyle.SOLID);

        // 7. Resaltar Fila de Próximo Pago en Magenta Claro 3 (#ead1dc)
        if (proximaFilaPago) {
            sheet.getRange(proximaFilaPago, 1, 1, 10).setBackground('#ead1dc');
        }

        // 8. Autoajustar anchos de columna
        sheet.setColumnWidth(1, 70);
        sheet.setColumnWidth(2, 100);
        sheet.setColumnWidth(3, 105);
        sheet.setColumnWidth(4, 90);
        sheet.setColumnWidth(5, 100);
        sheet.setColumnWidth(6, 100);
        sheet.setColumnWidth(7, 95);
        sheet.setColumnWidth(8, 95);
        sheet.setColumnWidth(9, 95);
        sheet.setColumnWidth(10, 100);

        // 9. REGLA ESTRICTA DE REGISTRO EN PRECIOS:
        // Solo registrar en Precios!S:T si hay una COMPRA REAL (tenencia en Cartera > 0 y != 100).
        let tenenciaRealEnCartera = 0;
        const hojaCartera = ss.getSheetByName(HOJAS.CARTERA);
        if (hojaCartera && (tickerPesos || tickerUSD)) {
            const lastRowCartera = hojaCartera.getLastRow();
            if (lastRowCartera >= 2) {
                const dataCartera = hojaCartera.getRange(2, 1, lastRowCartera - 1, 3).getValues();
                for (let r = 0; r < dataCartera.length; r++) {
                    const tickCartera = String(dataCartera[r][0] || '').trim().toUpperCase();
                    if (tickCartera === tickerPesos || tickCartera === tickerUSD) {
                        const cant = Number(dataCartera[r][2]) || 0;
                        if (cant > 0 && cant !== 100) {
                            tenenciaRealEnCartera = cant;
                        }
                        break;
                    }
                }
            }
        }

        let msgPrecios = '';
        if (config.registrarEnPrecios && tickerPesos) {
            if (tenenciaRealEnCartera > 0 && tenenciaRealEnCartera !== 100) {
                const hojaPrecios = ss.getSheetByName(HOJAS.PRECIOS);
                if (hojaPrecios) {
                    const maxRows = hojaPrecios.getMaxRows();
                    const rangoST = hojaPrecios.getRange(2, 19, maxRows - 1, 2).getValues();
                    let filaLibre = -1;
                    let yaExiste = false;

                    for (let r = 0; r < rangoST.length; r++) {
                        const pTicker = String(rangoST[r][0] || '').trim().toUpperCase();
                        const dTicker = String(rangoST[r][1] || '').trim().toUpperCase();

                        if (dTicker === tickerUSD || pTicker === tickerPesos) {
                            yaExiste = true;
                            break;
                        }
                        if (filaLibre === -1 && pTicker === '' && dTicker === '') {
                            filaLibre = r + 2;
                        }
                    }

                    if (!yaExiste && filaLibre !== -1) {
                        hojaPrecios.getRange(filaLibre, 19).setValue(tickerPesos);
                        hojaPrecios.getRange(filaLibre, 20).setValue(tickerUSD);
                        msgPrecios = ' (Registrado en Precios fila ' + filaLibre + ')';
                    }
                }
            } else {
                console.log('No se registra en Precios!S:T porque la tenencia es 100 (teórica/no comprado real).');
            }
        }

        return { success: true, message: 'Hoja ' + tickerUSD + ' creada exitosamente con ' + numFilas + ' cupones.' + msgPrecios };

    } catch (err) {
        console.error('Error en crearHojaBonoBackend: ' + err.message);
        return { success: false, message: err.message };
    }
}
