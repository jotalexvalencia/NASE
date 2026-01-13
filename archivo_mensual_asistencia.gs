// ===================================================================
// 📂 archivo_mensual_asistencia.gs – Archivo Histórico Mensual (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Archivo Automático de Nómina.
 * @description Este archivo automatiza la generación de copias de seguridad
 * de la hoja de nómina "Asistencia_SinValores" al final de cada mes.
 *
 * @workflow
 * - 🔁 **Trigger Automático:** Se ejecuta el día 1 de cada mes a las 12:00 PM.
 * - 📅 **Target:** Archiva los datos del *mes anterior* completo.
 * Ejemplo: Al ejecutarse el 1 de Febrero, el archivo dirá "Enero".
 * - 📁 **Ubicación:** Genera un Spreadsheet independiente en la carpeta:
 * "Archivos Asistencia Mensual NASE" dentro de Google Drive.
 *
 * @constraints
 * - ⛔ NO limpia la hoja original (la limpieza la hace el módulo bimestral).
 * - ✅ Aplica formato HH:mm:ss para evitar el error de visualización 1899.
 *
 * @author NASE Team
 * @version 1.2 (Corrección de Formato y Documentación Extendida)
 */

// ===================================================================
// 1. INSTALACIÓN DE DISPARADOR (TRIGGER)
// ===================================================================

/**
 * @summary Instala el disparador mensual de archivo.
 * @description Configura la ejecución recurrente para asegurar que cada mes
 * se genere un respaldo sin intervención humana.
 * * @schedule Día 1 de cada mes a las 12:00 PM (Mediodía).
 */
function instalarTriggersAsistenciaMensual() {
  ensureTimeTrigger("generarArchivoMensualAsistencia", function () {
    ScriptApp.newTrigger("generarArchivoMensualAsistencia")
      .timeBased()
      .onMonthDay(1) // Ejecución mensual el primer día
      .atHour(12)    // 12:00 PM
      .create();
  });
  Logger.log("✅ Trigger mensual Asistencia_SinValores instalado satisfactoriamente.");
}

// ===================================================================
// 2. LÓGICA DE ARCHIVO
// ===================================================================

/**
 * @summary Genera el archivo histórico consolidado del mes anterior.
 * @description Proceso técnico de 6 pasos:
 * 1. Validación de la hoja de origen "Asistencia_SinValores".
 * 2. Cálculo dinámico del nombre del mes anterior (Locale es-ES).
 * 3. Creación de un nuevo archivo de Google Sheets en Drive.
 * 4. Clonación de la hoja completa con formatos y fórmulas mediante .copyTo().
 * 5. Aplicación de NumberFormat "HH:mm:ss" para corregir la visualización de horas.
 * 6. Remoción de hojas residuales (Hoja 1) en el archivo de destino.
 */
function generarArchivoMensualAsistencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Asistencia_SinValores');
  
  // Validar existencia de datos antes de proceder.
  if (!hoja) {
    Logger.log("❌ Error: No se encontró la hoja origen para el proceso de archivo.");
    return;
  }

  // -----------------------------------------------------------
  // 1. CONTEXTO TEMPORAL (Determinación de Mes y Año)
  // -----------------------------------------------------------
  const ahora = new Date();
  // Se resta 1 al mes actual para obtener el periodo vencido.
  const mesAnterior = new Date(ahora.getFullYear(), ahora.getMonth() - 1, 1);
  const nombreMes = mesAnterior.toLocaleString('es-ES', { month: 'long', year: 'numeric' });
  const nombreArchivo = `Asistencia_${nombreMes.replace(' ', '_')}`;

  // -----------------------------------------------------------
  // 2. GESTIÓN DE CARPETAS EN DRIVE
  // -----------------------------------------------------------
  const folder = obtenerOCrearCarpeta('Archivos Asistencia Mensual NASE');

  // -----------------------------------------------------------
  // 3. CREACIÓN DEL RECURSO (Spreadsheet)
  // -----------------------------------------------------------
  const archivo = SpreadsheetApp.create(nombreArchivo);
  DriveApp.getFileById(archivo.getId()).moveTo(folder);

  // -----------------------------------------------------------
  // 4. COPIADO DE DATOS ESTRUCTURADOS
  // -----------------------------------------------------------
  // .copyTo() es el método más seguro para mantener la fidelidad de los datos.
  const hojaCopia = hoja.copyTo(archivo);
  // Se limita el nombre de la pestaña por restricciones de longitud de Sheets.
  hojaCopia.setName('Asistencia_' + nombreMes.substring(0, 15));

  // -----------------------------------------------------------
  // 5. NORMALIZACIÓN DE FORMATOS (HH:mm:ss)
  // -----------------------------------------------------------
  // Previene que las horas se transformen en fechas de 1899 al ser copiadas.
  const ultimaFila = hojaCopia.getLastRow();
  if (ultimaFila > 1) {
    hojaCopia.getRange(2, 1, ultimaFila - 1, hojaCopia.getLastColumn())
             .setNumberFormat("HH:mm:ss");
  }

  // -----------------------------------------------------------
  // 6. DEPURACIÓN DEL ARCHIVO DESTINO
  // -----------------------------------------------------------
  // SpreadsheetApp.create() siempre incluye una "Hoja 1". Procedemos a eliminarla
  // para que el archivo histórico solo contenga la información relevante.
  const hojas = archivo.getSheets();
  hojas.forEach(h => {
    if (h.getName() !== hojaCopia.getName()) {
      archivo.deleteSheet(h);
    }
  });

  Logger.log(`✅ Consolidado histórico generado con éxito: ${nombreArchivo}`);
}

// ===================================================================
// 3. UTILIDADES DE INFRAESTRUCTURA (DRIVE API)
// ===================================================================

/**
 * @summary Busca una carpeta en la raíz de Drive. Si es inexistente, la crea.
 * @param {String} nombre - Nombre descriptivo de la carpeta.
 * @returns {Folder} El objeto carpeta de Google Drive listo para su uso.
 */
function obtenerOCrearCarpeta(nombre) {
  const folders = DriveApp.getFoldersByName(nombre);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(nombre);
}
