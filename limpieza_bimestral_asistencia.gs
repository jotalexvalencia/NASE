// ===================================================================
// 🧹 limpieza_bimestral_asistencia.gs – Limpieza Automática (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo de Limpieza Automática (Ciclo Bimestral).
 * @description Gestiona la limpieza masiva de la hoja "Asistencia_SinValores"
 *              para evitar que el archivo de Google Sheets crezca indefinidamente.
 * 
 * @logic
 *   - ⚡ **Trigger:** Se programa para ejecutarse cada 2 meses.
 *       Específicamente el Día 1 de los meses impares (Enero, Marzo, Mayo, Julio, Septiembre, Noviembre)
 *       a las 15:00 (3:00 PM).
 *   - 🗑️ **Acción de Limpieza:** Borra todo el contenido de las filas de datos (dejando el encabezado).
 *   - 🛡️ **Seguridad (Sin Respaldo Local):**
 *       Este archivo NO crea respaldos dentro del Spreadsheet actual.
 *       Se asume que `archivo_mensual_asistencia.gs` (que corre el día 1 a las 12:00 PM)
 *       ya ha creado una copia de seguridad en Google Drive.
 *       Esto asegura que los datos del mes anterior no se pierdan antes de limpiar.
 *   - 🔁 **Ciclo:**
 *       Enero (Limpia) -> Febrero (No limpia) -> Marzo (Limpia) -> ...
 *
 * @dependencies
 *   - `install_triggers.gs` (Función `ensureTimeTrigger`).
 *   - `archivo_mensual_asistencia.gs` (Debe ejecutarse 3 horas antes para respaldar).
 *
 * @author NASE Team
 * @version 1.0
 */

// ===================================================================
// 1. INSTALACIÓN DEL DISPARADOR (TRIGGER)
// ===================================================================

/**
 * @summary Instala el disparador bimestral de limpieza.
 * @description Función de instalación (manual o al desplegar).
 *              Utiliza `ensureTimeTrigger` para verificar si ya existe.
 * 
 * @schedule
 *   - Día del mes: 1 (Primero de cada mes impar).
 *   - Hora: 15 (3:00 PM).
 *   - Frecuencia: Mensual, pero la función interna tiene un filtro de meses impares.
 */
function instalarTriggersLimpiezaBimestral() {
  // Wrapper de seguridad para crear el trigger
  ensureTimeTrigger("limpiarAsistenciaBimestral", function () {
    ScriptApp.newTrigger("limpiarAsistenciaBimestral")
      .timeBased()
      .onMonthDay(1) // Día 1
      .atHour(15)    // A las 15:00 (3:00 PM)
      .create();
  });
  Logger.log("✅ Trigger bimestral limpieza Asistencia_SinValores instalado.");
}

// ===================================================================
// 2. LÓGICA DE LIMPIEZA (Ciclo Bimestral)
// ===================================================================

/**
 * @summary Limpia la hoja de asistencia si corresponde al mes.
 * @description Función principal que se ejecuta automáticamente por el Trigger.
 *              Realiza lo siguiente:
 *   1. Obtiene la fecha actual del sistema.
 *   2. Verifica si el mes es impar (Enero, Marzo, Mayo, Julio, Septiembre, Noviembre).
 *   3. Si es impar, limpia la hoja "Asistencia_SinValores".
 *   4. Muestra un Toast en la hoja y un mensaje en Log.
 * 
 * @safety
 *   - Al ser bimestral (Cada 2 meses), el archivo permanece limpio por dos meses.
 *   - Se recomienda que el archivo mensual (`archivo_mensual_asistencia.gs`) corra
 *     siempre el día 1 a las 12:00 PM, 3 horas ANTES de esta limpieza, para respaldar.
 * 
 * @note NO crea respaldo interno. El respaldo es el archivo mensual en Drive.
 */
function limpiarAsistenciaBimestral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asistencia_SinValores");
  
  // Validación básica: Si no existe hoja o está vacía, no hacer nada
  if (!hoja || hoja.getLastRow() <= 1) return;

  const hoy = new Date();
  const mes = hoy.getMonth(); // 0=Enero, 1=Febrero, ... 11=Diciembre

  // -----------------------------------------------------------------
  // 1. FILTRO DE FRECUENCIA (Solo meses impares)
  // -----------------------------------------------------------------
  // Operación Modulo (% 2):
  // Si mes es 0 (Enero) -> Par (False) -> No ejecuta. 
  // Si mes es 1 (Febrero) -> Par (False) -> No ejecuta.
  // Si mes es 2 (Marzo) -> Impar (True) -> Ejecuta.
  // NOTA: Según la lógica bimestral, si queremos limpiar en Marzo (mes 2), es impar.
  // Si el ciclo es Ene/Feb -> Limpia Marzo, esto coincide con `mes % 2 !== 0`.
  // (Verificar si la intención es limpiar Ene, Mar, Mayo, etc. que son indices 0, 2, 4... PARES en 0-indexing, 
  //  pero IMPARES en fecha real. Enero=0 (Par), Marzo=2 (Par). 
  //  CORRECCIÓN DE LÓGICA: Si el prompt dice "Cada 2 meses", y "Dia 1 impar", 
  //  implica meses impares del calendario: Enero, Marzo, Mayo, Julio, Septiembre, Noviembre.
  //  En 0-indexing: Enero(0), Marzo(2), Mayo(4) son PARES.
  //  La lógica `if (mes % 2 !== 0)` ejecutará en Febrero(1), Abril(3)... (Impares).
  //  DEBO MANTENER LA LÓGICA DEL PROMPT O CORREGIRLA SEGÚN "MES IMPAR DEL CALENDARIO"?
  //  El prompt dice: "Se ejecuta cada 2 meses (día 1 impar)". Enero, Marzo, Mayo son meses 1, 3, 5.
  //  En 0-indexing son 0, 2, 4.
  //  La condición `if (mes % 2 !== 0)` en el código original ejecuta en Febrero, Abril... (Impares).
  //  Para ejecutar en Enero, Marzo, Mayo, la condición debe ser `if (mes % 2 === 0)`.
  //  SIN EMBARGO, NO DEBO CAMBIAR LÓGICA. Documentaré lo que hace el código tal cual.
  
  // ⚠️ IMPORTANTE: El código original usa `if (mes % 2 !== 0)`.
  // Esto significa que se ejecutará en Febrero, Abril, Junio, Agosto, Octubre, Diciembre (Meses impares del calendario).
  // Mantendré la documentación fiel al código.
  if (mes % 2 !== 0) return; 

  // -----------------------------------------------------------------
  // 2. ACCIÓN DE LIMPIEZA (Borrar Filas)
  // -----------------------------------------------------------------
  
  // ✅ Solo limpiar, sin crear respaldo interno
  // El respaldo se confía al archivo mensual generado anteriormente
  const lastRow = hoja.getLastRow();
  
  if (lastRow > 1) {
    // Borra desde la fila 2 hasta la última fila, todas las columnas
    // Mantiene los encabezados (fila 1)
    hoja.getRange(2, 1, lastRow - 1, hoja.getLastColumn()).clearContent();
  }

  // -----------------------------------------------------------------
  // 3. FEEDBACK VISUAL (Toast y Log)
  // -----------------------------------------------------------------
  
  // Mostrar mensaje en la hoja para el usuario
  SpreadsheetApp.getActive().toast(
    `✅ Limpieza bimestral Asistencia_SinValores completada.`,
    "Limpieza completada",
    8 // Segundos visibles
  );

  Logger.log(`✅ Limpieza bimestral Asistencia_SinValores completada. Sin respaldo interno.`);
}
