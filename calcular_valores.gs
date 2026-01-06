// ============================================================
// 💰 calcular_valores.gs – Formateo y Limpieza de Asistencia (NASE 2026)
// ------------------------------------------------------------
/**
 * @summary Módulo de Formateo y Mantenimiento de la Hoja de Asistencia.
 * @description 
 * Este archivo ha evolucionado para eliminar la lógica compleja de nómina.
 * Actualmente se especializa en la presentación de datos:
 * 
 * ⚠️ CAMBIOS RECIENTES (2025/2026):
 * - ❌ ELIMINADO: Cálculo de nómina (Salario * Recargo).
 * - ❌ ELIMINADO: Inserción de columnas de dinero ($).
 * - ✅ NUEVA FUNCIÓN: Aplicar formato numérico estándar a columnas de TIEMPO.
 * - ✅ NUEVA FUNCIÓN: Eliminar rastros de columnas monetarias viejas.
 *
 * 🎨 Lógica de Formateo:
 * - Identifica columnas de Tiempo (Ej: "Total Horas", "Horas Diurnas").
 * - Aplica formato `#,##0.00` (miles con punto, decimales con coma).
 * - Busca y destruye columnas obsoletas (Ej: "Valor Diurno Domingo/Festivo").
 *
 * @author NASE Team
 * @version 2.0 (Clean Code - Sin Nómina)
 */

// ============================================================
// 1. FUNCIÓN PRINCIPAL
// ============================================================

/**
 * @summary Procesa y limpia la hoja "Asistencia_SinValores".
 * @description Función pública que se ejecuta manual o por trigger.
 *              Realiza dos acciones principales:
 *              1. **Formatear Tiempo:** Asegura que las horas se vean legibles (#,##0.00).
 *              2. **Eliminar Rastros Monetarios:** Busca y borra columnas de valores ($) que
 *                 pertenecían a versiones antiguas del sistema, para evitar confusión.
 * 
 * @note Esta función NO calcula salarios. El cálculo se ha movido a herramientas
 *       externas o procesos administrativos fuera de este script.
 */
function agregarValoresAsistencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asistencia_SinValores");
  
  if (!hoja) {
    Logger.log("❌ No existe la hoja 'Asistencia_SinValores'. Ejecuta primero la generación de turnos.");
    SpreadsheetApp.getUi().alert("⚠️ No existe la hoja 'Asistencia_SinValores'.");
    return;
  }

  const ultimaFila = hoja.getLastRow();
  
  // Verificar que haya datos (excluyendo encabezados)
  if (ultimaFila < 2) {
    Logger.log("⚠️ No hay datos para procesar en Asistencia_SinValores.");
    return;
  }

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

  // -----------------------------------------------------------
  // 1. IDENTIFICACIÓN DE COLUMNAS DE TIEMPO
  // -----------------------------------------------------------
  const findHeaderIndex = nombre =>
    encabezados.findIndex(h => String(h).trim().toLowerCase() === nombre.trim().toLowerCase());

  // Índices de columnas relacionadas con TIEMPO (Horas trabajadas)
  const colHorasTotales = findHeaderIndex("Total Horas");
  const colHorasDiurnas = findHeaderIndex("Horas Diurnas");
  const colHorasNocturnas = findHeaderIndex("Horas Nocturnas Normales");
  const colDomDiurnas = findHeaderIndex("Horas Diurnas Domingo/Festivo");
  const colDomNocturnas = findHeaderIndex("Horas Nocturnas Domingo/Festivo");

  // Lista de índices de columnas de TIEMPO encontradas
  const indicesTiempo = [colHorasTotales, colHorasDiurnas, colHorasNocturnas, colDomDiurnas, colDomNocturnas];

  // -----------------------------------------------------------
  // 2. ACCIÓN: APLICAR FORMATO NUMÉRICO
  // -----------------------------------------------------------
  // Aplica formato #,##0.00 (miles con punto, decimales con coma)
  // a todas las columnas de tiempo detectadas desde la fila 2 hasta el final.
  indicesTiempo.forEach(idx => {
    if (idx !== -1) {
      // El formato "#,##0.00" permite mostrar números grandes (ej: 123.45,67)
      // sin símbolo de moneda, ideal para sumas de horas o tarifas base.
      hoja.getRange(2, idx + 1, ultimaFila, 1).setNumberFormat("#,##0.00");
    }
  });

  // -----------------------------------------------------------
  // 3. ACCIÓN: LIMPIEZA DE COLUMNAS MONETARIAS ANTIGUAS
  // -----------------------------------------------------------
  // Lista de nombres de columnas que deben ser eliminadas si existen.
  // Estas columnas correspondían a cálculos internos de nómina de versiones anteriores.
  const colsMonetariasABorrar = [
    "Valor Diurno Domingo/Festivo",
    "Valor Nocturno Día Ordinario",
    "Valor Nocturno Domingo/Festivo",
    "Total Valores" // Suma monetaria antigua
  ];

  // Recorremos de derecha a izquierda (última columna hacia la primera)
  // Esto es CRÍTICO porque al borrar una columna, los índices de las columnas a su derecha cambian.
  let borradas = 0;
  
  for (let i = encabezados.length - 1; i >= 0; i--) {
    const headerName = String(encabezados[i]).trim();
    
    // Si el encabezado coincide con alguno de los nombres prohibidos
    if (colsMonetariasABorrar.includes(headerName)) {
      // Eliminamos la columna (índice i + 1, porque el array es 0-based)
      hoja.deleteColumn(i + 1);
      borradas++;
      Logger.log(`🗑️ Columna monetaria eliminada: ${headerName}`);
    }
  }

  // -----------------------------------------------------------
  // 4. RESULTADOS Y LOGGING
  // -----------------------------------------------------------
  Logger.log(`✅ Proceso finalizado. Formato aplicado a ${indicesTiempo.length} columnas de tiempo. ${borradas} columnas monetarias eliminadas.`);
  
  // Feedback visual al usuario si ejecuta manualmente
  try {
    SpreadsheetApp.getActive().toast("Formateo completado.", "Asistencia", 5);
  } catch(e) {
    // Ignorar error si no hay UI disponible (ej: ejecución por Trigger)
  }
}
