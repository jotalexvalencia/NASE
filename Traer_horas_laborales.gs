// ============================================================
// ⏱️ Traer_horas_laborales.gs – Sincronización de Horas Laborales
// ------------------------------------------------------------
/**
 * @summary Módulo de Sincronización de Datos Laborales (Base Operativa vs NASE).
 * @description Esta función conecta el sistema NASE actual con el libro externo 
 *              de "Base Operativa" (RRHH) para trae la configuración de horas 
 *              laborales semanales de cada empleado.
 * 
 * @features
 * - 🔗 **Conexión Externa:** Abre otro libro de Google Sheets usando su ID (`ID_BASE_OPERATIVA`).
 * - 📂 **Gestión de Columnas:** Verifica si existe la columna "Horas Laborales por Semana"
 *   en la hoja de asistencia. Si no existe, la inserta automáticamente.
 * - 🧠 **Mapa de Memoria:** Construye un mapa temporal (Objeto JS) para procesar 
 *   la base operativa. Esto permite soportar cientos de registros sin usar 
 *   memoria excesiva de script.
 * - 🔄 **Lógica de Último Registro:** Si un empleado tiene múltiples entradas 
 *   (ej: contrato antiguo, contrato nuevo, actualización de horas), el script
 *   selecciona automáticamente aquella con la "Fecha de Ingreso" más reciente.
 * - ✅ **Actualización Masiva:** Busca coincidencias de cédula en la hoja de asistencia
 *   y actualiza la columna de horas en una sola operación (Batch update).
 *
 * @dependencies
 * - Libro externo: `1bU-lyiQzczid62n8timgUguW6UxC3qZN8Vehnn26zdY` (Base Operativa).
 * - Hoja externa: `BASE OPERATIVA`.
 * - Hoja interna: `Asistencia_SinValores`.
 * - Columnas requeridas en Base Operativa: `DOCUMENTO DE IDENTIDAD`, 
 *   `HORAS LABORALES POR SEMANA`, `FECHA DE INGRESO`.
 *
 * @author NASE Team
 * @version 1.2 (Optimizado con Mapa de Memoria)
 */

// ======================================================================
// FUNCIÓN PRINCIPAL
// ======================================================================

/**
 * @summary Sincroniza las "Horas Laborales por Semana" desde RRHH a Asistencia.
 * @description 
 * 1. Abre el libro de "Base Operativa".
 * 2. Lee todos los empleados y busca el registro más reciente por fecha de ingreso.
 * 3. Lee la hoja "Asistencia_SinValores".
 * 4. Crea la columna "Horas Laborales por Semana" si falta.
 * 5. Cruza los datos por Cédula y actualiza las horas.
 * 
 * @note Esta función es ideal para ejecutarse manualmente cuando se actualiza
 *       la nómina o cada vez que se cambia el esquema de contratos.
 * 
 * @returns {void} Escribe en `Logger` y muestra alerta de éxito.
 */
function insertarHorasLaboralesPorCedula() {
  // -----------------------------------------------------------
  // 1. CONFIGURACIÓN Y APERTURA DE LIBROS
  // -----------------------------------------------------------
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAsistencia = ss.getSheetByName("Asistencia_SinValores");
  
  // Validar hoja destino
  if (!hojaAsistencia) {
    throw new Error("❌ No se encontró la hoja 'Asistencia_SinValores'.");
  }

  // ID del Libro de RRHH (Base Operativa) - Fuente de datos
  const ID_BASE_OPERATIVA = "1bU-lyiQzczid62n8timgUguW6UxC3qZN8Vehnn26zdY";
  const libroBase = SpreadsheetApp.openById(ID_BASE_OPERATIVA);
  const hojaBase = libroBase.getSheetByName("BASE OPERATIVA");

  // Validar hoja origen
  if (!hojaBase) {
    throw new Error("❌ No se encontró la hoja 'BASE OPERATIVA' en la Base Operativa.");
  }

  // -----------------------------------------------------------
  // 2. PREPARACIÓN DE HOJA DESTINO (Asistencia)
  // -----------------------------------------------------------
  
  // Leer encabezados de asistencia
  const headersAsistencia = hojaAsistencia.getRange(1, 1, 1, hojaAsistencia.getLastColumn()).getValues()[0];
  
  // Buscar índice de columna 'Cédula'
  const colCedulaAsistencia = headersAsistencia.findIndex(h => String(h).trim().toLowerCase() === "cédula") + 1;
  
  if (colCedulaAsistencia === 0) {
    throw new Error("⚠️ No se encontró la columna 'Cédula' en Asistencia_SinValores.");
  }

  // Definir nombre de la columna a insertar/actualizar
  const nombreColumnaNueva = "Horas Laborales por Semana";
  
  // Buscar índice de la nueva columna (para verificar si ya existe)
  let colNueva = headersAsistencia.findIndex(h => String(h).trim() === nombreColumnaNueva) + 1;

  // -----------------------------------------------------------
  // 3. GESTIÓN DE COLUMNAS (Crear si falta)
  // -----------------------------------------------------------
  
  // Si la columna NO existe, insertarla justo después de la columna 'Cédula'
  if (colNueva === 0) {
    hojaAsistencia.insertColumnAfter(colCedulaAsistencia);
    // Escribir el encabezado en la nueva columna creada
    hojaAsistencia.getRange(1, colCedulaAsistencia + 1).setValue(nombreColumnaNueva);
    // Actualizar el índice local a la nueva posición
    colNueva = colCedulaAsistencia + 1;
  }

  // -----------------------------------------------------------
  // 4. LECTURA Y PROCESAMIENTO DE DATOS ORIGEN (Base Operativa)
  // -----------------------------------------------------------
  
  // Leer toda la base operativa en memoria
  const dataBase = hojaBase.getDataRange().getValues();
  const headersBase = dataBase[0];
  
  // Normalizar encabezados base a mayúsculas para búsqueda robusta
  const headersBaseUpper = headersBase.map(h => (h || "").toString().trim().toUpperCase());

  // Buscar índices de columnas clave en la Base Operativa
  const idxCedulaBase = headersBaseUpper.indexOf("DOCUMENTO DE IDENTIDAD");
  const idxHorasBase = headersBaseUpper.indexOf("HORAS LABORALES POR SEMANA");
  const idxFechaBase = headersBaseUpper.indexOf("FECHA DE INGRESO");

  // Validar que existan las columnas necesarias
  if ([idxCedulaBase, idxHorasBase, idxFechaBase].includes(-1)) {
    throw new Error("⚠️ Faltan columnas requeridas: 'DOCUMENTO DE IDENTIDAD', 'HORAS LABORALES POR SEMANA', 'FECHA DE INGRESO'.");
  }

  // -----------------------------------------------------------
  // 5. CREAR MAPA DE MEMORIA { Cédula -> { Horas, Fecha } }
  // -----------------------------------------------------------
  /**
   * Lógica Crucial: La Base Operativa puede tener varios registros por empleado
   * (ej: contratos anteriores, reingresos, ajustes de horas).
   * 
   * Necesitamos QUITAR UNO: El más reciente según 'FECHA DE INGRESO'.
   * 
   * MapaHoras almacenará: { "12345678": { horas: 40, fecha: DateObj } }
   */
  const mapaHoras = {};

  for (let i = 1; i < dataBase.length; i++) {
    const fila = dataBase[i];
    const cedula = String(fila[idxCedulaBase]).replace(/\D/g, "").trim();
    
    if (!cedula) continue; // Omitir cédulas vacías

    const horas = fila[idxHorasBase];
    const fechaIngreso = fila[idxFechaBase];
    
    // Validar fecha
    if (!fechaIngreso) continue;

    // Convertir fecha a objeto Date para comparación
    const fecha = fechaIngreso instanceof Date ? fechaIngreso : new Date(fechaIngreso);
    
    // Validar objeto fecha
    if (!fecha || isNaN(fecha)) continue;

    // ALGORITMO DE ACTUALIZACIÓN:
    // 1. Si la cédula NO está en el mapa -> Agregar.
    // 2. Si la cédula SÍ está en el mapa -> Verificar fecha.
    //    Si la fecha actual es MAYOR a la fecha guardada -> Actualizar.
    // Esto asegura que se tomen los contratos o actualizaciones más recientes.
    if (!mapaHoras[cedula] || fecha > mapaHoras[cedula].fecha) {
      mapaHoras[cedula] = { 
        horas: horas, 
        fecha: fecha 
      };
    }
  }

  // -----------------------------------------------------------
  // 6. ACTUALIZACIÓN DE HOJA DESTINO (Asistencia)
  // -----------------------------------------------------------
  const ultimaFila = hojaAsistencia.getLastRow();
  
  // No hacer nada si no hay datos de asistencia
  if (ultimaFila < 2) return Logger.log("⚠️ No hay registros en Asistencia_SinValores.");

  // Leer todas las cédulas de la hoja de asistencia (columna A)
  const cedulas = hojaAsistencia.getRange(2, colCedulaAsistencia, ultimaFila - 1, 1).getValues();
  
  // Crear array de valores para escribir (1 sola columna con N filas)
  // Iterar sobre las cédulas de asistencia y buscar su valor de horas en el Mapa
  const valores = cedulas.map(([cedula]) => {
    const c = String(cedula || "").replace(/\D/g, "").trim();
    // Retornar el valor del mapa, o vacío si la cédula no está en la base operativa
    return [mapaHoras[c] ? mapaHoras[c].horas : ""];
  });

  // Escribir todos los valores de golpe en la hoja de asistencia
  hojaAsistencia.getRange(2, colNueva, valores.length, 1).setValues(valores);

  Logger.log(`✅ Columna '${nombreColumnaNueva}' actualizada correctamente (${valores.length} filas procesadas).`);
  
  // Feedback visual para el usuario si ejecuta manualmente
  SpreadsheetApp.getActive().toast("✅ Horas laborales sincronizadas correctamente.", "Base Operativa", 5);
}
