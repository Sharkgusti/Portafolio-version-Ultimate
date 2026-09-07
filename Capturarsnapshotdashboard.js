// =================================================================================
// ===  CAPTURARSNAPSHOTDASHBOARD.GS — Script standalone, de SOLO LECTURA       ===
// ===  Paso previo al 3.4. Guarda una "foto" del JSON completo que devuelve   ===
// ===  generarDatosMaestros() ANTES de migrarla, para comparar campo por      ===
// ===  campo después del cambio.                                             ===
// ===                                                                          ===
// ===  No modifica nada del sistema real. Único efecto secundario: crea un    ===
// ===  archivo de texto en tu Google Drive (carpeta raíz).                    ===
// =================================================================================

function CAPTURAR_SNAPSHOT_DASHBOARD_ANTES() {
    const ui = SpreadsheetApp.getUi();

    let datos;
    try {
        datos = generarDatosMaestros();
    } catch (err) {
        ui.alert('❌ Error al ejecutar generarDatosMaestros(): ' + err.message);
        return;
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const nombreArchivo = 'SNAPSHOT_DASHBOARD_ANTES_' + timestamp + '.json';

    const json = JSON.stringify(datos, null, 2);
    const archivo = DriveApp.createFile(nombreArchivo, json, MimeType.PLAIN_TEXT);

    const mensaje = '✅ Snapshot guardado en tu Drive:\n\n"' + nombreArchivo + '"\n\n' +
        'Tamaño: ' + Math.round(json.length / 1024) + ' KB\n\n' +
        'Guardá el nombre de este archivo — lo vamos a necesitar para comparar ' +
        'después de migrar generarDatosMaestros() en el Paso 3.4.\n\n' +
        'Link directo: ' + archivo.getUrl();

    ui.alert(mensaje);
    console.log(mensaje);
    console.log('ID del archivo (por si hace falta): ' + archivo.getId());
}