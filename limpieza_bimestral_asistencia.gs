// ===================================================================
// 🧹 limpieza_bimestral_asistencia.gs – Limpieza Automática (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Limpieza Automática (Ciclo Bimestral).
 * @description Gestiona la limpieza masiva de la hoja "Asistencia_SinValores"
 * para evitar que el archivo de Google Sheets crezca indefinidamente.
 * * @logic
 * - ⚡ **Trigger:** Se programa para ejecutarse cada 2 meses.
 * Específicamente el Día 1 de los meses que cumplan la condición lógica
 * a las 15:00 (3:00 PM).
 * - 🗑️ **Acción de Limpieza:** Borra todo el contenido de las filas de datos (dejando el encabezado).
 * - 🛡️ **Seguridad (Sin Respaldo Local):**
 * Este archivo NO crea respaldos dentro del Spreadsheet actual.
 * Se asume que `archivo_mensual_asistencia.gs` (que corre el día 1 a las 12:00 PM)
 * ya ha creado una copia de seguridad en Google Drive.
 * Esto asegura que los datos del mes anterior no se pierdan antes de limpiar.
 * - 🔁 **Ciclo:**
 * Se ejecuta en meses alternos según la validación de módulo (mes % 2).
 *
 * @dependencies
 * - `install_triggers.gs` (Función `ensureTimeTrigger`).
 * - `archivo_mensual_asistencia.gs` (Debe ejecutarse 3 horas antes para respaldar).
 *
 * @author NASE Team
 * @version 1.1 (Lógica Bimestral Validada)
 */

// ===================================================================
// 1. INSTALACIÓN DEL DISPARADOR (TRIGGER)
// ===================================================================

/**
 * @summary Instala el disparador bimestral de limpieza.
 * @description Función de instalación (manual o al desplegar).
 * Utiliza `ensureTimeTrigger` para verificar si ya existe un disparador previo
 * y evitar duplicaciones innecesarias de tareas programadas.
 * * @schedule
 * - Día del mes: 1.
 * - Hora: 15 (3:00 PM).
 * - Frecuencia: Mensual, filtrada internamente por la función lógica.
 */
function instalarTriggersLimpiezaBimestral() {
  // Wrapper de seguridad para crear el trigger si no existe
  ensureTimeTrigger("limpiarAsistenciaBimestral", function () {
    ScriptApp.newTrigger("limpiarAsistenciaBimestral")
      .timeBased()
      .onMonthDay(1) // Día 1 de cada mes
      .atHour(15)    // A las 15:00 (3:00 PM)
      .create();
  });
  Logger.log("✅ Trigger bimestral limpieza Asistencia_SinValores instalado.");
}

// ===================================================================
// 2. LÓGICA DE LIMPIEZA (Ciclo Bimestral)
// ===================================================================

/**
 * @summary Limpia la hoja de asistencia si corresponde al mes actual.
 * @description Función principal que se ejecuta automáticamente por el Trigger.
 * Realiza las siguientes validaciones y acciones:
 * 1. Obtiene la hoja "Asistencia_SinValores".
 * 2. Verifica si el mes actual cumple la condición de salto (módulo 2).
 * 3. Si la condición se cumple, procede a borrar el contenido desde la fila 2.
 * 4. Lanza un aviso visual (Toast) para informar a los usuarios activos.
 * * @safety
 * - Al ser bimestral, la hoja acumula datos de 60 días antes de ser vaciada.
 * - Es CRÍTICO que el respaldo mensual haya ocurrido horas antes.
 * * @note NO genera archivos. Solo vacía la hoja para mantener el rendimiento.
 */
function limpiarAsistenciaBimestral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asistencia_SinValores");
  
  // Validación de existencia: Si la hoja no existe o solo tiene encabezados, salir.
  if (!hoja || hoja.getLastRow() <= 1) return;

  const hoy = new Date();
  const mes = hoy.getMonth(); // 0=Enero, 1=Febrero...

  // -----------------------------------------------------------------
  // 1. FILTRO DE FRECUENCIA (Lógica de meses alternos)
  // -----------------------------------------------------------------
  // Se mantiene la lógica !== 0 para respetar el ciclo definido por el usuario.
  if (mes % 2 !== 0) return; 

  // -----------------------------------------------------------------
  // 2. ACCIÓN DE LIMPIEZA (Borrar Filas preservando formato)
  // -----------------------------------------------------------------
  const lastRow = hoja.getLastRow();
  
  if (lastRow > 1) {
    // .clearContent() borra los datos pero mantiene colores, bordes y formatos de celda.
    // Esto es ideal para que AppSheet siga viendo una estructura limpia.
    hoja.getRange(2, 1, lastRow - 1, hoja.getLastColumn()).clearContent();
  }

  // -----------------------------------------------------------------
  // 3. FEEDBACK VISUAL Y AUDITORÍA
  // -----------------------------------------------------------------
  SpreadsheetApp.getActive().toast(
    `✅ Limpieza bimestral Asistencia_SinValores completada.`,
    "Mantenimiento NASE",
    8
  );

  Logger.log(`✅ Operación exitosa: Se limpiaron ${lastRow - 1} filas de la hoja de asistencia.`);
} 
