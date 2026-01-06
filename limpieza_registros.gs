// ============================================================
// 🧹 limpieza_registros.gs – Filtrado y Generación de Reportes (NASE 2026)
// ------------------------------------------------------------
/**
 * @summary Módulo de Filtrado y Gestión de Reportes Seguros.
 * @description Filtra registros de la hoja "Respuestas" basándose en reglas
 *              temporales y los copia a una hoja "Filtrado".
 * 
 * @safety
 *   - 🛡️ NO BORRA datos: La hoja principal ("Respuestas") permanece intacta.
 *   - 📄 COPIA A REPORTES: Genera una hoja nueva o sobrescribe "Filtrado" con los datos.
 *
 * @criteria (Reglas de Negocio)
 *   - 🔹 Criterio 1 (Nocturno/Mes Anterior): Registros del último día del mes anterior
 *     que ocurrieron entre las 18:00 y las 22:00 (Turnos de cierre de mes).
 *   - 🔹 Criterio 2 (Mes Actual): TODOS los registros del mes actual en curso.
 * 
 * @author NASE Team
 * @version 1.3 (Versión Corregida con Manejo Seguro de Fechas)
 */

// ======================================================================
// FUNCIÓN PRINCIPAL: Filtro y Generación
// ======================================================================

/**
 * @summary Genera reporte de asistencia filtrando por fechas.
 * @description Ejecuta la lógica de doble criterio para extraer registros relevantes
 *              para reportes de nómina o auditoría.
 * 
 * @workflow
 *   1. Abre las hojas "Respuestas" (Origen) y "Filtrado" (Destino).
 *   2. Calcula dinámicamente las fechas del mes anterior y actual.
 *   3. Lee todos los registros y agrupa por Cédula (para tratar por empleado).
 *   4. Aplica Filtro 1: Registros del último día del mes anterior entre 18:00-22:00.
 *   5. Aplica Filtro 2: Todos los registros del mes actual.
 *   6. Escribe el resultado final en la hoja "Filtrado".
 * 
 * @requires Hoja "Respuestas" con columnas: Cédula, Centro, Fecha Entrada, Hora Entrada, Fecha Salida, Hora Salida.
 */
function filtrarRegistrosUltimoDiaMesAnteriorYMesActual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // -----------------------------------------------------------
  // 1. CONFIGURACIÓN DE HOJAS (Origen y Destino)
  // -----------------------------------------------------------
  let hojaDestino = ss.getSheetByName("Filtrado");
  
  // Si no existe la hoja destino, crearla
  if (!hojaDestino) hojaDestino = ss.insertSheet("Filtrado");

  // Hoja origen (Base de datos principal)
  const hojaOrigen = ss.getSheetByName("Respuestas");
  if (!hojaOrigen) {
    ui.alert("❌ No se encontró la hoja 'Respuestas'.");
    return;
  }

  // -----------------------------------------------------------
  // 2. OBTENER INDICES DE COLUMNAS (Manejo Dinámico)
  // -----------------------------------------------------------
  const datos = hojaOrigen.getDataRange().getValues();
  
  if (datos.length < 2) {
    ui.alert("⚠️ La hoja 'Respuestas' está vacía o solo tiene encabezados.");
    return;
  }

  // Buscar índices por nombre (insensible a mayúsculas/espacios)
  const encabezados = datos[0];
  
  const idxCedula = encabezados.indexOf("Cédula");
  const idxCentro = encabezados.indexOf("Centro");       // Útil para reporte
  const idxFechaEnt = encabezados.indexOf("Fecha Entrada");
  const idxHoraEnt = encabezados.indexOf("Hora Entrada");
  const idxFechaSal = encabezados.indexOf("Fecha Salida");
  const idxHoraSal = encabezados.indexOf("Hora Salida");
  const idxDentroSal = encabezados.indexOf("Dentro Salida"); // Útil para reporte
  const idxNombre = encabezados.indexOf("Nombre");             // Útil para reporte

  // Validar columnas esenciales
  if (idxCedula === -1 || idxFechaEnt === -1 || idxHoraEnt === -1) {
    ui.alert("❌ No se encontraron las columnas necesarias ('Cédula', 'Fecha Entrada', 'Hora Entrada').");
    return;
  }

  // -----------------------------------------------------------
  // 3. CÁLCULO DE FECHAS DEL SISTEMA
  // -----------------------------------------------------------
  const hoy = new Date();
  const mesActual = hoy.getMonth(); // 0 = Enero
  const anioActual = hoy.getFullYear();
  
  // Cálculo del último día del mes anterior (Truco de día 0)
  // Si hoy es 1 de Noviembre (mes 10), new Date(2024, 10, 0) es 31 de Octubre.
  const ultimoDiaMesAnterior = new Date(anioActual, mesActual, 0); 
  const diaUltimoMesAnterior = ultimoDiaMesAnterior.getDate();
  const mesAnterior = ultimoDiaMesAnterior.getMonth();
  const anioMesAnterior = ultimoDiaMesAnterior.getFullYear();

  // -----------------------------------------------------------
  // 4. AGRUPACIÓN DE REGISTROS POR CÉDULA
  // -----------------------------------------------------------
  /**
   * Mapa para acumular filas por empleado.
   * Estructura: { "12345678": [filaObjeto, filaObjeto, ...], ... }
   */
  const mapaCedulas = {};

  // Recorremos toda la hoja de Respuestas
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const cedula = fila[idxCedula];
    
    // Ignorar filas sin cédula
    if (!cedula) continue;

    // Reconstrucción robusta de fecha/hora de entrada
    const fechaRaw = fila[idxFechaEnt];
    const horaRaw = fila[idxHoraEnt];
    
    let fecha = null;

    // ✅ FIX: Convertir a String antes de split para evitar error
    const fechaStr = String(fechaRaw || '').trim();
    const horaStr = String(horaRaw || '').trim();

    if (fechaStr && horaStr) {
       // Formato esperado: dd/mm/yyyy HH:mm
       const parts = fechaStr.split('/');
       // Parseo manual a ISO (YYYY-MM-DDTHH:mm) para evitar errores de zona horaria
       if (parts.length === 3) {
         fecha = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaStr}`);
       }
    }

    // Solo procesar filas con fechas válidas
    if (cedula && fecha && !isNaN(fecha.getTime())) {
      // Si es la primera vez que vemos esta cédula, creamos el array
      if (!mapaCedulas[cedula]) mapaCedulas[cedula] = [];
      
      // Guardamos la fila entera en el array de la cédula
      mapaCedulas[cedula].push(fila); 
    }
  }

  // Array para acumular las filas finales que pasen el filtro
  const filasFinales = [];

  // -----------------------------------------------------------
  // 5. LÓGICA DE FILTRADO (Por Empleado)
  // -----------------------------------------------------------
  
  // Iteramos sobre cada empleado en el mapa
  for (const cedula in mapaCedulas) {
    const filasCedula = mapaCedulas[cedula];

    // ---------------------------------------------------------
    // 🔹 CRITERIO 1: Último día del mes anterior (Horario Nocturno)
    // ---------------------------------------------------------
    const registrosUltimoDia = filasCedula.filter(fila => {
      const fechaRaw = fila[idxFechaEnt];
      const horaRaw = fila[idxHoraEnt];
      
      let fecha = null;
      const fechaStr = String(fechaRaw || '').trim();
      const horaStr = String(horaRaw || '').trim();

      // Parsear fecha igual que arriba
      if (fechaStr && horaStr) {
         const parts = fechaStr.split('/');
         if (parts.length === 3) fecha = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaStr}`);
      }
      
      // Si no hay fecha válida, descartar
      if (!fecha) return false;
      
      // Extraer hora
      const h = parseInt(horaStr.split(':')[0], 10);
      
      // Comprobar rango nocturno (18:00 a 22:00)
      // Nota: Se usa hora exacta, no se considera minutos para el rango del criterio
      const esNoche = (h >= 18 && h <= 22);
      
      // Verificar si coincide con el "Último día del mes anterior"
      return (
        fecha.getFullYear() === anioMesAnterior &&
        fecha.getMonth() === mesAnterior &&
        fecha.getDate() === diaUltimoMesAnterior &&
        esNoche
      );
    });

    // Si se encontraron registros de cierre nocturno, agregarlos
    if (registrosUltimoDia.length > 0) {
      // Usamos Spread Operator para agregar el array completo
      filasFinales.push(...registrosUltimoDia);
    }

    // ---------------------------------------------------------
    // 🔹 CRITERIO 2: Registros del mes actual (Completo)
    // ---------------------------------------------------------
    const registrosMesActual = filasCedula.filter(fila => {
      const fechaRaw = fila[idxFechaEnt];
      const horaRaw = fila[idxHoraEnt];
      
      let fecha = null;
      const fechaStr = String(fechaRaw || '').trim();
      const horaStr = String(horaRaw || '').trim();

      if (fechaStr && horaStr) {
         const parts = fechaStr.split('/');
         if (parts.length === 3) fecha = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T${horaStr}`);
      }

      if (!fecha) return false;
      
      // Verificar si pertenece al año y mes actual
      return fecha.getFullYear() === anioActual && fecha.getMonth() === mesActual;
    });

    // Agregar todos los registros del mes actual
    filasFinales.push(...registrosMesActual);
  }

  // -----------------------------------------------------------
  // 6. ESCRITURA DE RESULTADO EN HOJA DESTINO
  // -----------------------------------------------------------
  
  // Si no se encontró nada, avisar
  if (filasFinales.length === 0) {
    ui.alert("❌ No se encontraron registros válidos para el filtro actual.");
    return;
  }

  // Limpiar contenido de la hoja "Filtrado" para evitar datos basura de corridas anteriores
  hojaDestino.clearContents();
  
  // Escribir encabezados originales
  hojaDestino.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  
  // Escribir filas filtradas
  hojaDestino.getRange(2, 1, filasFinales.length, encabezados.length).setValues(filasFinales);

  // -----------------------------------------------------------
  // 7. COPIA DE SEGURIDAD (Opcional / Comentado)
  // -----------------------------------------------------------
  
  // Lista de nombres de meses para nombres de archivos
  const nombreMeses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const nombreMesAnterior = nombreMeses[mesAnterior];
  
  // Generar nombre para archivo histórico (ej: registro_octubre_2025)
  const nombreRespaldo = `registro_${nombreMesAnterior}_${anioMesAnterior}`;
  
  // Lógica para crear una copia en Drive como respaldo (Deshabilitada por defecto)
  // const hojaExistente = ss.getSheetByName(nombreRespaldo);
  // if (hojaExistente) ss.deleteSheet(nombreRespaldo); // Borrar respaldo viejo si existe
  // hojaOrigen.copyTo(ss).setName(nombreRespaldo); // Crear nuevo respaldo

  // Mostrar notificación de éxito
  SpreadsheetApp.getActive().toast(
    `✅ Se generaron ${filasFinales.length} registros en la hoja "Filtrado".`,
    "Reporte Generado",
    5
  );
}
