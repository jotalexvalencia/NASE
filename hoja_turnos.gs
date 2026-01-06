// ============================================================
// 📘 hoja_turnos.gs – Generador de Asistencia Completo (NASE 2026)
// ------------------------------------------------------------
/**
 * @summary Módulo Central de Cálculo de Nómina y Asistencia.
 * @description Orquesta la generación de la hoja "Asistencia_SinValores"
 *              basándose en los registros de Entrada/Salida ("Respuestas").
 *
 * 🏗️ ARQUITECTURA DEL SISTEMA:
 * - ⚡ **Procesamiento por Lotes (Batching):** Lee hasta 10,000 filas de "Respuestas"
 *   para procesarlas y escribir los resultados, luego usa un Trigger para continuar.
 *   Esto evita el error de "Tiempo de ejecución superado" (6 mins) de Google.
 * - 🧬 **Disparadores (Triggers):** Usa `ScriptApp.newTrigger().after()` para
 *   crear una cadena de ejecución que procesa todos los datos automáticamente.
 * - 🧠 **Bloqueo (Lock):** Usa `LockService` para evitar que dos usuarios o procesos
 *   generen asistencia al mismo tiempo y corrompan los datos.
 * - 🧮 **Algoritmo de Horas:** Calcula minuto a minuto (o en chunks) basándose en
 *   configuraciones dinámicas (Inicio/Fín Nocturno) y calendario de Festivos.
 *
 * 📊 SALIDA (Hoja "Asistencia_SinValores"):
 * - Genera una fila por cada turno trabajado por cada empleado.
 * - Desglosa horas en 4 categorías:
 *   1. Horas Diurnas Normales.
 *   2. Horas Nocturnas Normales.
 *   3. Horas Diurnas Domingo/Festivo.
 *   4. Horas Nocturnas Domingo/Festivo.
 *
 * ✅ CORRECCIONES APLICADAS (2026-01-06):
 * - Columnas alineadas correctamente (14 columnas).
 * - Hora Inicio y Hora Salida muestran valores reales del turno.
 * - Tipo de Día detectado correctamente (Normal/Domingo/Festivo).
 * - Festivos colombianos calculados con Ley Emiliani.
 *
 * @dependencies
 * - Code.gs (buscarEmpleadoPorCedula).
 * - ConfigHorarios.gs (obtenerConfiguracionHorarios).
 *
 * @author NASE Team
 * @version 4.0 (Corregida - Columnas y Festivos)
 */

// ===================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// ===================================================================

/** @summary Nombre de la hoja origen con los registros de entrada/salida. */
const HOJA_ORIGEN = 'Respuestas';

/** @summary Nombre de la hoja destino donde se genera la asistencia. */
const HOJA_DESTINO = 'Asistencia_SinValores';

/** @summary Tamaño del lote de filas a procesar por ejecución (evita timeout de 6 min). */
const TAMANO_LOTE = 10000;

/** @summary Nombre de la función handler del trigger de continuación. */
const ASIS_TRIGGER_HANDLER = 'continuarProcesoAsistencia';

/** @summary Propiedad donde se guarda la fila actual del proceso. */
const ASIS_PROP_LOTE_INICIO = 'ASIS_LOTE_INICIO';

/** @summary Propiedad flag "0" o "1" para saber si el proceso está corriendo. */
const ASIS_PROP_EN_CURSO = 'ASIS_EN_CURSO';

// ===================================================================
// 2. UTILIDADES DE NOTIFICACIÓN (UI Segura)
// ===================================================================

/**
 * @summary Notificador seguro (UI / Toast / Logger).
 * @description Intenta mostrar una alerta (si hay UI, ej: ejecución manual).
 *              Si falla (ej: trigger), intenta mostrar un Toast.
 *              Si falla todo, registra en Logger.
 *
 * @param {String} message - El mensaje a mostrar.
 * @param {String} title - Título de la alerta/toast.
 */
function _asisNotify_(message, title) {
  var t = title || 'Asistencia NASE';
  var msg = String(message || '');

  // 1. Intentar UI Alert (Solo funciona en ejecución manual desde Sheets)
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(t, msg, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  } catch (e) {}

  // 2. Intentar UI Toast (Puede funcionar en algunas ejecuciones)
  try {
    SpreadsheetApp.getActive().toast(msg, t, 8);
    return;
  } catch (e) {}

  // 3. Fallback Logger (Para Triggers o Errores de permisos)
  Logger.log('[' + t + '] ' + msg);
}

/**
 * @summary Verifica si hay una Interfaz de Usuario activa.
 * @returns {Boolean} True si se ejecuta desde el Editor o la Hoja.
 */
function _asisHasUi_() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * @summary Muestra un Toast seguro (No rompe si no hay UI).
 * @param {String} message - Mensaje a mostrar.
 * @param {String} title - Título del toast.
 * @param {Number} seconds - Duración en segundos.
 */
function _asisToastSafe_(message, title, seconds) {
  try {
    SpreadsheetApp.getActive().toast(String(message || ''), String(title || ''), Number(seconds || 5));
  } catch (e) {}
}

// ===================================================================
// 3. FUNCIÓN PRINCIPAL (Orquestador)
// ===================================================================

/**
 * @summary Inicia el proceso de generación de asistencia.
 * @description Esta es la función que se invoca manual o por Trigger.
 *              - Verifica bloqueos (Lock).
 *              - Prepara la hoja destino (la limpia).
 *              - Inicia el bucle de lotes.
 *
 * @returns {void} Muestra alerta o logs de progreso.
 */
function generarTablaAsistenciaSinValores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDatos = ss.getSheetByName(HOJA_ORIGEN);

  // Validar hoja origen
  if (!hojaDatos) {
    _asisNotify_('Error: No se encontró la hoja "' + HOJA_ORIGEN + '".', 'Error');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();

  // Si no se puede obtener el bloqueo, significa que otro proceso está corriendo.
  if (!lock.tryLock(3000)) {
    _asisNotify_('Otro proceso está usando el generador. Intenta nuevamente en unos segundos.', 'Sistema ocupado');
    return;
  }

  try {
    // Si la propiedad 'ASIS_EN_CURSO' vale '1', significa que ya se inició
    if (props.getProperty(ASIS_PROP_EN_CURSO) === '1') {
      _asisNotify_('Ya hay un proceso en curso. Continuará automáticamente.', 'Proceso en curso');
      _asisEnsureUniqueTrigger_(1000);
      return;
    }

    // Preparar hoja destino (Crear o Limpiar)
    var hojaSalida = ss.getSheetByName(HOJA_DESTINO);
    if (hojaSalida) {
      hojaSalida.clearContents();
    } else {
      hojaSalida = ss.insertSheet(HOJA_DESTINO);
    }

    // Warm-up: Intentar actualizar caché de empleados si la función existe
    if (typeof actualizarCacheEmpleados === 'function') {
      try {
        actualizarCacheEmpleados();
      } catch (e) {
        Logger.log('ASIS: actualizarCacheEmpleados falló: ' + e);
      }
    }

    // Inicializar propiedades de control
    props.deleteProperty(ASIS_PROP_LOTE_INICIO);
    props.setProperty(ASIS_PROP_LOTE_INICIO, '2');
    props.setProperty(ASIS_PROP_EN_CURSO, '1');

    // Mensaje inicial seguro
    _asisNotify_('Proceso iniciado. Se ejecutará por lotes hasta finalizar.', 'Asistencia');

    // Iniciar el primer lote
    _procesarLotesAsistencia();

  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// 4. PROCESAMIENTO POR LOTES (Batching)
// ===================================================================

/**
 * @summary Procesa un lote de datos y programa el siguiente lote.
 * @description Lee filas desde `ASIS_PROP_LOTE_INICIO` hasta `TAMANO_LOTE`.
 *              Las procesa, las escribe y programa el trigger `continuarProcesoAsistencia`.
 */
function _procesarLotesAsistencia() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var origen = ss.getSheetByName(HOJA_ORIGEN);
  var destino = ss.getSheetByName(HOJA_DESTINO);

  // Obtener fila actual donde estamos
  var inicio = parseInt(props.getProperty(ASIS_PROP_LOTE_INICIO) || '2', 10);
  var total = origen.getLastRow();

  // Si ya llegamos al final, finalizar proceso
  if (inicio > total) {
    _finalizarProcesoAsistencia();
    return;
  }

  // Calcular fin del lote (No exceder el total de filas)
  var fin = Math.min(inicio + TAMANO_LOTE - 1, total);

  // Leer datos de "Respuestas"
  var rango = origen.getRange(inicio, 1, fin - inicio + 1, origen.getLastColumn());
  var datos = rango.getValues();
  var tz = ss.getSpreadsheetTimeZone();

  // Procesar datos brutos -> Filas de asistencia
  var datosProcesados = _procesarDatosAsistencia(datos, tz);

  // Escribir resultados en "Asistencia_SinValores"
  if (datosProcesados.length > 0) {
    var inicioEscritura = destino.getLastRow() > 0 ? destino.getLastRow() + 1 : 1;

    // Si es la primera fila escrita, poner encabezados
    if (inicioEscritura === 1) {
      var encabezados = _generarEncabezados();
      destino.getRange(inicioEscritura, 1, 1, encabezados.length).setValues([encabezados]);
      inicioEscritura++;
    }

    // Escribir lote de filas
    destino.getRange(inicioEscritura, 1, datosProcesados.length, datosProcesados[0].length)
      .setValues(datosProcesados);
  }

  // Guardar progreso para el siguiente ciclo
  props.setProperty(ASIS_PROP_LOTE_INICIO, String(fin + 1));

  // Calcular porcentaje y log
  var porcentaje = ((fin / total) * 100).toFixed(1);
  Logger.log('📦 Procesadas ' + fin + '/' + total + ' filas (' + porcentaje + '%)');

  // Feedback visual
  _asisToastSafe_('Procesando: ' + porcentaje + '% completado', 'Generando tabla de asistencia', 5);

  // Programar el siguiente lote (trigger en 1000ms)
  _asisEnsureUniqueTrigger_(1000);
}

/**
 * @summary Handler del Trigger para continuar el proceso.
 * @description Llamada automáticamente por el trigger. Limpia el trigger anterior
 *              para evitar duplicados y reanuda el procesamiento.
 */
function continuarProcesoAsistencia() {
  var lock = LockService.getScriptLock();

  // Intentar bloqueo por 30 segundos. Si hay otro proceso, salir.
  if (!lock.tryLock(30000)) return;

  try {
    _asisClearTriggers_();
    _procesarLotesAsistencia();
  } finally {
    lock.releaseLock();
  }
}

/**
 * @summary Finaliza el proceso de asistencia.
 * @description Limpia propiedades, borra triggers y muestra mensaje final.
 */
function _finalizarProcesoAsistencia() {
  var props = PropertiesService.getScriptProperties();

  // Borrar propiedades de estado
  props.deleteProperty(ASIS_PROP_LOTE_INICIO);
  props.deleteProperty(ASIS_PROP_EN_CURSO);

  // Borrar cualquier trigger pendiente
  _asisClearTriggers_();

  Logger.log('✅ Proceso completado: hoja Asistencia_SinValores lista.');

  // Verificación de datos generados
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_DESTINO);
  var ultimaFila = hoja.getLastRow();

  Logger.log('Filas creadas en Asistencia_SinValores: ' + (ultimaFila - 1));

  if (ultimaFila <= 1) {
    Logger.log('⚠️ Advertencia: No se crearon registros. Verificar logs anteriores.');
  }

  // Notificación final
  _asisNotify_('✅ Proceso completado. La hoja Asistencia_SinValores ha sido generada.', 'Asistencia');
}

// ===================================================================
// 5. GESTIÓN DE TRIGGERS (Control de Automatización)
// ===================================================================

/**
 * @summary Busca triggers por nombre de función handler.
 * @param {String} handler - Nombre de la función handler.
 * @returns {Array} Array de triggers encontrados.
 * @private
 */
function _asisGetTriggersByHandler_(handler) {
  return ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === handler;
  });
}

/**
 * @summary Borra todos los triggers manejados por ASIS.
 * @private
 */
function _asisClearTriggers_() {
  var ts = _asisGetTriggersByHandler_(ASIS_TRIGGER_HANDLER);
  ts.forEach(function(t) {
    try {
      ScriptApp.deleteTrigger(t);
    } catch (e) {}
  });
}

/**
 * @summary Asegura que exista UN SOLO trigger de continuación.
 * @description Si no existe, lo crea. Si existen varios, los borra.
 * @param {Number} delayMs - Milisegundos de espera antes de disparar.
 * @private
 */
function _asisEnsureUniqueTrigger_(delayMs) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return;

  try {
    var existentes = _asisGetTriggersByHandler_(ASIS_TRIGGER_HANDLER);

    // Si ya existe 1, no hacer nada
    if (existentes.length > 0) return;

    // Si no existe, crearlo
    ScriptApp.newTrigger(ASIS_TRIGGER_HANDLER)
      .timeBased()
      .after(Math.max(300, Number(delayMs) || 500))
      .create();

  } catch (e) {
    Logger.log('ASIS: no se pudo crear el trigger: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// 6. GENERACIÓN DE ENCABEZADOS
// ===================================================================

/**
 * @summary Genera el array de encabezados para la hoja de asistencia.
 * @description ✅ 14 columnas que coinciden con los datos generados.
 * @returns {Array} Array de strings con los nombres de columnas.
 * @private
 */
function _generarEncabezados() {
  return [
    "Cédula",                           // 0
    "Nombre Empleado",                  // 1
    "Centro",                           // 2
    "Rango de Fecha",                   // 3: Formato "dd/MM/yyyy HH:mm - dd/MM/yyyy HH:mm"
    "Fecha",                            // 4: Solo fecha de inicio
    "Hora Inicio",                      // 5
    "Hora Salida",                      // 6
    "Tipo Día Inicio",                  // 7
    "Tipo Día Fin",                     // 8
    "Horas Trabajadas",                 // 9
    "Horas Diurnas Normales",           // 10
    "Horas Nocturnas Normales",         // 11
    "Horas Diurnas Domingo/Festivo",    // 12
    "Horas Nocturnas Domingo/Festivo"   // 13
  ];
}

// ===================================================================
// 7. MOTOR DE PROCESAMIENTO DE DATOS (Lógica Core)
// ===================================================================

/**
 * @summary Procesa datos crudos y genera filas de asistencia calculadas.
 * @description ✅ VERSIÓN CORREGIDA:
 *              - Genera 14 columnas alineadas con encabezados.
 *              - Usa hora real del turno (no medianoche).
 *              - Tipo de día detectado correctamente.
 *
 * @param {Array} datos - Array de filas crudas de la hoja "Respuestas".
 * @param {String} tz - Zona horaria (Ej: "America/Bogota").
 * @returns {Array} Array de Arrays listo para escribir en la hoja.
 */
function _procesarDatosAsistencia(datos, tz) {
  // Mapeo de índices (Debe coincidir con RESP_HEADERS de Code.gs)
  var IDX = {
    CEDULA: 0,
    CENTRO: 1,
    FECHA_ENT: 14,
    HORA_ENT: 15,
    FECHA_SAL: 17,
    HORA_SAL: 18
  };

  // Paso 1: Filtrar y Normalizar Filas de Entrada/Salida
  var registros = [];

  for (var i = 0; i < datos.length; i++) {
    var r = datos[i];

    // Validar fila válida
    if (!r || r.length < 3 || !r[IDX.CEDULA]) continue;

    var cedula = String(r[IDX.CEDULA] || '').trim();
    var centroReal = String(r[IDX.CENTRO] || '').trim();

    // Extraer Fechas y Horas
    var fechaEntrada = r[IDX.FECHA_ENT];
    var horaEntrada = r[IDX.HORA_ENT];
    var fechaSalida = r[IDX.FECHA_SAL];
    var horaSalida = r[IDX.HORA_SAL];

    // Validar que existan fechas y horas
    if (!fechaEntrada || !horaEntrada) continue;

    // Convertir a objetos Date
    var tsEntrada = _parsearFechaHora(fechaEntrada, horaEntrada, tz);
    var tsSalida = null;

    // Si hay fecha Salida, la parseamos
    if (fechaSalida && horaSalida) {
      tsSalida = _parsearFechaHora(fechaSalida, horaSalida, tz);
    }

    // Validar secuencia lógica (Salida debe ser posterior a Entrada)
    if (!tsEntrada || !tsSalida || tsSalida <= tsEntrada) {
      continue;
    }

    registros.push({
      cedula: cedula,
      centroReal: centroReal,
      inicio: tsEntrada,
      fin: tsSalida,
      completo: true
    });
  }

  if (registros.length === 0) {
    Logger.log('⚠️ No se procesaron registros válidos (faltan fechas o hay inconsistencias).');
    return [];
  }

  // Paso 2: Generar filas de asistencia
  var cacheFestivos = {};
  var filas = [];

  for (var j = 0; j < registros.length; j++) {
    var turno = registros[j];
    var ced = turno.cedula;
    var centro = turno.centroReal;

    // Buscar datos del empleado
    var emp = { ok: false, nombre: 'Sin Nombre' };
    if (typeof buscarEmpleadoPorCedula === 'function') {
      emp = buscarEmpleadoPorCedula(ced);
    }
    var nombre = emp.ok ? emp.nombre : 'NO ENCONTRADO';

    // ✅ CORRECCIÓN: Crear RANGO DE FECHA completo
    var fechaInicioStr = Utilities.formatDate(turno.inicio, tz, "dd/MM/yyyy");
    var horaInicioStr = Utilities.formatDate(turno.inicio, tz, "HH:mm");
    var fechaFinStr = Utilities.formatDate(turno.fin, tz, "dd/MM/yyyy");
    var horaFinStr = Utilities.formatDate(turno.fin, tz, "HH:mm");

    var rangoCompleto = fechaInicioStr + " " + horaInicioStr + " - " + fechaFinStr + " " + horaFinStr;

    // ✅ CORRECCIÓN: Detectar tipo de día CORRECTAMENTE
    var tipoDiaInicio = _esFestivo(turno.inicio, cacheFestivos);
    var tipoDiaFin = _esFestivo(turno.fin, cacheFestivos);

    // Calcular horas trabajadas
    var horas = _calcularHorasPorTipo([{ inicio: turno.inicio, fin: turno.fin }], cacheFestivos);

    // ✅ CORRECCIÓN: Construir fila con 14 columnas
    var fila = [
      ced,                // 0: Cédula
      nombre,             // 1: Nombre Empleado
      centro,             // 2: Centro
      rangoCompleto,      // 3: Rango de Fecha
      fechaInicioStr,     // 4: Fecha (solo fecha de inicio)
      horaInicioStr,      // 5: Hora Inicio (CORREGIDO)
      horaFinStr,         // 6: Hora Salida (CORREGIDO)
      tipoDiaInicio,      // 7: Tipo Día Inicio
      tipoDiaFin,         // 8: Tipo Día Fin
      horas.total,        // 9: Horas Trabajadas
      horas.normalesDia,  // 10: Horas Diurnas Normales
      horas.normalesNoc,  // 11: Horas Nocturnas Normales
      horas.festivosDia,  // 12: Horas Diurnas Domingo/Festivo
      horas.festivosNoc   // 13: Horas Nocturnas Domingo/Festivo
    ];

    filas.push(fila);
  }

  return filas;
}

// ===================================================================
// 8. UTILIDADES DE PARSEO DE FECHA/HORA
// ===================================================================

/**
 * @summary Parsea una fecha y hora (String u Objeto) a un objeto Date.
 * @description Maneja formato dd/mm/yyyy HH:mm o Date Objects nativos de Sheets.
 *
 * @param {String|Date} fechaVal - Valor de la celda fecha.
 * @param {String|Date} horaVal - Valor de la celda hora.
 * @param {String} tz - Zona horaria ("America/Bogota").
 * @returns {Date|null} Objeto Date parseado.
 * @private
 */
function _parsearFechaHora(fechaVal, horaVal, tz) {
  try {
    var fecha = null;

    // Si es Objeto Date nativo de Google Sheets
    if (fechaVal instanceof Date) {
      fecha = new Date(fechaVal);
    } else if (typeof fechaVal === 'string') {
      var strFecha = fechaVal.trim();

      // Si contiene formato en inglés (Mon, Tue, etc.), intentar parsear directo
      if (strFecha.includes('Mon') || strFecha.includes('Tue') || strFecha.includes('Wed') ||
          strFecha.includes('Thu') || strFecha.includes('Fri') || strFecha.includes('Sat') ||
          strFecha.includes('Sun') || strFecha.includes('GMT')) {
        fecha = new Date(strFecha);
      } else {
        // Asumir formato dd/MM/yyyy
        var parts = strFecha.split('/');
        if (parts.length === 3) {
          fecha = new Date(parts[2] + '-' + parts[1] + '-' + parts[0]);
        }
      }
    }

    if (!fecha || isNaN(fecha.getTime())) {
      return null;
    }

    var hh = 0, mm = 0, ss = 0;

    // Si hora es Objeto Date
    if (horaVal instanceof Date) {
      hh = horaVal.getHours();
      mm = horaVal.getMinutes();
      ss = horaVal.getSeconds();
    } else if (typeof horaVal === 'string') {
      var strHora = horaVal.trim();

      // Si contiene formato largo con 1899 o GMT, extraer solo HH:mm:ss
      var matchHora = strHora.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (matchHora) {
        hh = parseInt(matchHora[1], 10);
        mm = parseInt(matchHora[2], 10);
        ss = matchHora[3] ? parseInt(matchHora[3], 10) : 0;
      }
    }

    if (isNaN(hh) || isNaN(mm)) {
      return null;
    }

    // Establecer horas en el objeto Date
    fecha.setHours(hh, mm, ss, 0);
    return fecha;

  } catch (e) {
    Logger.log('❌ Error parseando fecha/hora: ' + fechaVal + ' ' + horaVal + ' - ' + e.message);
    return null;
  }
}

/**
 * @summary Trunca una fecha a minutos exactos (elimina segundos y milisegundos).
 * @param {Date} fecha - Fecha a truncar.
 * @returns {Date} Fecha truncada.
 * @private
 */
function _truncarMinutos(fecha) {
  if (!fecha || !(fecha instanceof Date)) return null;
  return new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate(), fecha.getHours(), fecha.getMinutes(), 0, 0);
}

/**
 * @summary Formatea una fecha como string YYYY-MM-DD.
 * @param {Number} anio - Año.
 * @param {Number} mes - Mes (1-12).
 * @param {Number} dia - Día.
 * @returns {String} Fecha formateada.
 * @private
 */
function _formatearFecha(anio, mes, dia) {
  return anio + '-' + String(mes).padStart(2, '0') + '-' + String(dia).padStart(2, '0');
}

// ===================================================================
// 9. DETECCIÓN DE TIPO DE DÍA (Normal/Domingo/Festivo)
// ===================================================================

/**
 * @summary Determina si una fecha es Normal, Domingo o Festivo.
 * @description ✅ VERSIÓN CORREGIDA: Retorna string descriptivo.
 *              Genera la lista de festivos del año si no existe en caché.
 *
 * @param {Date} fecha - Fecha a evaluar.
 * @param {Object} cacheFestivos - Objeto para guardar festivos calculados.
 * @returns {String} "Normal", "Domingo" o "Festivo".
 * @private
 */
function _esFestivo(fecha, cacheFestivos) {
  // Validación de entrada
  if (!(fecha instanceof Date) || isNaN(fecha.getTime())) {
    return "Normal";
  }

  var anio = fecha.getFullYear();

  // Generar festivos del año si no están en caché
  if (!cacheFestivos[anio]) {
    cacheFestivos[anio] = _generarFestivos(anio);
  }

  // Formatear fecha para buscar en Set
  var mes = fecha.getMonth() + 1;
  var dia = fecha.getDate();
  var clave = anio + '-' + String(mes).padStart(2, '0') + '-' + String(dia).padStart(2, '0');

  // ✅ Primero verificar si es festivo
  if (cacheFestivos[anio].has(clave)) {
    return "Festivo";
  }

  // ✅ Luego verificar si es domingo
  if (fecha.getDay() === 0) {
    return "Domingo";
  }

  // ✅ Si no es ninguno, es día normal
  return "Normal";
}

// ===================================================================
// 10. MOTOR DE CÁLCULO DE HORAS (Lógica de Negocio)
// ===================================================================

/**
 * @summary Calcula las horas trabajadas clasificadas por tipo.
 * @description ✅ VERSIÓN CORREGIDA: Usa comparación de strings para tipos de día.
 *              Este es el cerebro financiero de la nómina.
 *
 * @param {Array} intervalos - Array de objetos {inicio: Date, fin: Date}.
 * @param {Object} cacheFestivos - Caché de festivos calculados.
 * @returns {Object} { total, normalesDia, normalesNoc, festivosDia, festivosNoc }
 */
function _calcularHorasPorTipo(intervalos, cacheFestivos) {
  // Validar entrada
  if (!Array.isArray(intervalos)) {
    intervalos = [intervalos];
  }

  // Leer configuración dinámica de horarios nocturnos
  var config = { horaInicio: 21, horaFin: 6 }; // Default: 9PM a 6AM
  if (typeof obtenerConfiguracionHorarios === 'function') {
    try {
      config = obtenerConfiguracionHorarios();
    } catch (e) {
      Logger.log('Usando configuración nocturna por defecto');
    }
  }
  var horaInicioNoc = config.horaInicio || 21;
  var horaFinNoc = config.horaFin || 6;

  var normalesDia = 0;
  var normalesNoc = 0;
  var festivosDia = 0;
  var festivosNoc = 0;

  for (var idx = 0; idx < intervalos.length; idx++) {
    var intervalo = intervalos[idx];

    if (!intervalo || !intervalo.inicio || !intervalo.fin) continue;

    var inicio = _truncarMinutos(intervalo.inicio);
    var fin = _truncarMinutos(intervalo.fin);

    if (!inicio || !fin || isNaN(inicio.getTime()) || isNaN(fin.getTime()) || fin <= inicio) continue;

    var cursor = new Date(inicio);

    while (cursor < fin) {
      var cursorTime = cursor.getTime();
      var cursorHora = cursor.getHours();

      // Calcular siguiente punto de cambio
      var nextMidnight = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1, 0, 0, 0);

      var nextInicioNoc = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate(), horaInicioNoc, 0, 0);
      if (cursorTime >= nextInicioNoc.getTime()) {
        nextInicioNoc.setDate(nextInicioNoc.getDate() + 1);
      }

      var nextFinNoc = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate(), horaFinNoc, 0, 0);
      if (cursorTime >= nextFinNoc.getTime()) {
        nextFinNoc.setDate(nextFinNoc.getDate() + 1);
      }

      // Determinar siguiente límite
      var limites = [fin.getTime(), nextMidnight.getTime(), nextInicioNoc.getTime(), nextFinNoc.getTime()];
      var limitesValidos = [];
      for (var li = 0; li < limites.length; li++) {
        if (limites[li] > cursorTime) limitesValidos.push(limites[li]);
      }

      var siguienteTs = Math.min.apply(null, limitesValidos);

      if (!siguienteTs || siguienteTs <= cursorTime) {
        cursor = new Date(cursorTime + 60000);
        continue;
      }

      var siguiente = new Date(siguienteTs);

      // Calcular horas del chunk
      var delta = (siguiente.getTime() - cursor.getTime()) / (1000 * 60 * 60);

      // ✅ CORRECCIÓN: Determinar si es horario nocturno
      var esNocturna = false;
      if (horaInicioNoc > horaFinNoc) {
        // Nocturno cruza medianoche (ej: 21:00 - 06:00)
        esNocturna = (cursorHora >= horaInicioNoc || cursorHora < horaFinNoc);
      } else {
        // Nocturno no cruza medianoche
        esNocturna = (cursorHora >= horaInicioNoc && cursorHora < horaFinNoc);
      }

      // ✅ CORRECCIÓN: Comparar STRING con STRING
      var tipoDelDia = _esFestivo(cursor, cacheFestivos);

      // Clasificar y sumar
      if (tipoDelDia === "Normal") {
        if (esNocturna) {
          normalesNoc += delta;
        } else {
          normalesDia += delta;
        }
      } else {
        // "Domingo" o "Festivo" -> Ambos son recargo festivo
        if (esNocturna) {
          festivosNoc += delta;
        } else {
          festivosDia += delta;
        }
      }

      cursor = siguiente;
    }
  }

  return {
    total: Number((normalesDia + normalesNoc + festivosDia + festivosNoc).toFixed(2)),
    normalesDia: Number(normalesDia.toFixed(2)),
    normalesNoc: Number(normalesNoc.toFixed(2)),
    festivosDia: Number(festivosDia.toFixed(2)),
    festivosNoc: Number(festivosNoc.toFixed(2))
  };
}

// ===================================================================
// 11. LÓGICA DE FESTIVOS COLOMBIA (Cálculo de Fechas Móviles)
// ===================================================================

/**
 * @summary Calcula la fecha del Domingo de Pascua (Algoritmo de Meeus/Jones/Butcher).
 * @description Utilizado para calcular Semana Santa y festivos móviles.
 * @param {Number} anio - Año a calcular.
 * @returns {Date} Objeto Date con la fecha del Domingo de Pascua.
 * @private
 */
function _calcularDomingoPascua(anio) {
  var a = anio % 19;
  var b = Math.floor(anio / 100);
  var c = anio % 100;
  var d = Math.floor(b / 4);
  var e = b % 4;
  var f = Math.floor((b + 8) / 25);
  var g = Math.floor((b - f + 1) / 3);
  var h = (19 * a + b - d - g + 15) % 30;
  var i = Math.floor(c / 4);
  var k = c % 4;
  var l = (32 + 2 * e + 2 * i - h - k) % 7;
  var m = Math.floor((a + 11 * h + 22 * l) / 451);
  var mes = Math.floor((h + l - 7 * m + 114) / 31);
  var dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(anio, mes - 1, dia);
}

/**
 * @summary Genera el Set de festivos para Colombia.
 * @description ✅ VERSIÓN CORREGIDA con festivos colombianos correctos.
 *              - Festivos Fijos que NO se mueven.
 *              - Festivos con Ley Emiliani (se mueven al lunes siguiente).
 *              - Festivos móviles (Semana Santa, Corpus, etc.).
 *
 * @param {Number} anio - Año calendario.
 * @returns {Set<String>} Set de strings en formato "YYYY-MM-DD".
 */
function _generarFestivos(anio) {
  var festivos = new Set();

  // Función auxiliar: Mover al siguiente lunes (Ley Emiliani)
  var moverAlLunes = function(fecha) {
    var f = new Date(fecha);
    if (f.getDay() === 1) return f; // Ya es lunes

    // Calcular días para llegar al lunes
    var diasALunes = (8 - f.getDay()) % 7;
    if (diasALunes === 0) diasALunes = 7; // Si es lunes (resultado 0), ir al siguiente

    f.setDate(f.getDate() + diasALunes);
    return f;
  };

  // Función auxiliar: Agregar festivo al Set
  var agregarFestivo = function(fecha) {
    var y = fecha.getFullYear();
    var m = String(fecha.getMonth() + 1).padStart(2, '0');
    var d = String(fecha.getDate()).padStart(2, '0');
    festivos.add(y + '-' + m + '-' + d);
  };

  // ===================================================================
  // 1. FESTIVOS FIJOS (NO se mueven)
  // ===================================================================
  agregarFestivo(new Date(anio, 0, 1));   // 1 Enero - Año Nuevo
  agregarFestivo(new Date(anio, 4, 1));   // 1 Mayo - Día del Trabajo
  agregarFestivo(new Date(anio, 6, 20));  // 20 Julio - Independencia
  agregarFestivo(new Date(anio, 7, 7));   // 7 Agosto - Batalla de Boyacá
  agregarFestivo(new Date(anio, 11, 8));  // 8 Diciembre - Inmaculada Concepción
  agregarFestivo(new Date(anio, 11, 25)); // 25 Diciembre - Navidad

  // ===================================================================
  // 2. FESTIVOS CON LEY EMILIANI (Se mueven al lunes siguiente)
  // ===================================================================
  agregarFestivo(moverAlLunes(new Date(anio, 0, 6)));   // 6 Enero - Reyes Magos
  agregarFestivo(moverAlLunes(new Date(anio, 2, 19)));  // 19 Marzo - San José
  agregarFestivo(moverAlLunes(new Date(anio, 5, 29)));  // 29 Junio - San Pedro y San Pablo
  agregarFestivo(moverAlLunes(new Date(anio, 7, 15)));  // 15 Agosto - Asunción de la Virgen
  agregarFestivo(moverAlLunes(new Date(anio, 9, 12)));  // 12 Octubre - Día de la Raza
  agregarFestivo(moverAlLunes(new Date(anio, 10, 1)));  // 1 Noviembre - Todos los Santos
  agregarFestivo(moverAlLunes(new Date(anio, 10, 11))); // 11 Noviembre - Independencia de Cartagena

  // ===================================================================
  // 3. FESTIVOS MÓVILES (Basados en Pascua)
  // ===================================================================
  var domingoPascua = _calcularDomingoPascua(anio);

  // Jueves Santo (-3 días desde Pascua) - NO se mueve
  var juevesSanto = new Date(domingoPascua);
  juevesSanto.setDate(domingoPascua.getDate() - 3);
  agregarFestivo(juevesSanto);

  // Viernes Santo (-2 días desde Pascua) - NO se mueve
  var viernesSanto = new Date(domingoPascua);
  viernesSanto.setDate(domingoPascua.getDate() - 2);
  agregarFestivo(viernesSanto);

  // Ascensión del Señor (+39 días desde Pascua) - SE MUEVE al lunes
  var ascension = new Date(domingoPascua);
  ascension.setDate(domingoPascua.getDate() + 39);
  agregarFestivo(moverAlLunes(ascension));

  // Corpus Christi (+60 días desde Pascua) - SE MUEVE al lunes
  var corpus = new Date(domingoPascua);
  corpus.setDate(domingoPascua.getDate() + 60);
  agregarFestivo(moverAlLunes(corpus));

  // Sagrado Corazón (+68 días desde Pascua) - SE MUEVE al lunes
  var sagrado = new Date(domingoPascua);
  sagrado.setDate(domingoPascua.getDate() + 68);
  agregarFestivo(moverAlLunes(sagrado));

  return festivos;
}

// ===================================================================
// 12. FUNCIONES DE RESET MANUAL
// ===================================================================

/**
 * @summary Función manual para resetear el sistema si se queda colgado.
 * @description Limpia propiedades y triggers.
 */
function asistenciaResetManual() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(ASIS_PROP_LOTE_INICIO);
  props.deleteProperty(ASIS_PROP_EN_CURSO);
  _asisClearTriggers_();

  _asisNotify_('🔄 Generador de asistencia reseteado.', 'Asistencia');
}

/**
 * @summary Función pública para desbloquear el sistema manualmente.
 * @description Alias de asistenciaResetManual con mensaje UI.
 */
function resetearProcesoAsistencia() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(ASIS_PROP_LOTE_INICIO);
  props.deleteProperty(ASIS_PROP_EN_CURSO);
  _asisClearTriggers_();

  Logger.log('✅ Proceso de asistencia reseteado manualmente');

  try {
    SpreadsheetApp.getUi().alert('✅ Sistema desbloqueado. Puedes volver a generar la hoja.');
  } catch (e) {
    Logger.log('Sistema desbloqueado (sin UI disponible)');
  }
}

// ===================================================================
// 13. BACKEND PARA HTML (consulta/asistencia)
// ===================================================================

/**
 * @summary Obtiene datos de la hoja "Asistencia_SinValores" para mostrar en el HTML.
 * @description Lee la hoja generada, aplica filtros de fechas y devuelve JSON limpio.
 *
 * @param {Object} filtros - { fechaInicio: String, fechaFin: String }.
 * @returns {Object} { status: 'ok', registros: Array<{...}> }.
 */
function obtenerDataAsistencia(filtros) {
  filtros = filtros || {};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_DESTINO);

  if (!hoja) {
    return { status: 'error', message: 'La hoja Asistencia_SinValores no existe. Ejecuta primero el generador de turnos.' };
  }

  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return { status: 'ok', registros: [] };

  // Leer datos usando getDisplayValues para ver el valor como usuario lo ve
  var datos = hoja.getRange(2, 1, lastRow - 1, hoja.getLastColumn()).getDisplayValues();

  // Parsear filtros de fecha
  var fInicio = filtros.fechaInicio ? new Date(filtros.fechaInicio + 'T00:00:00') : null;
  var fFin = filtros.fechaFin ? new Date(filtros.fechaFin + 'T23:59:59') : null;

  var registros = [];

  for (var i = 0; i < datos.length; i++) {
    var fila = datos[i];

    // Fila[4] es la fecha para filtrar
    var fechaFila = null;
    if (fila[4]) {
      var parts = fila[4].split('/');
      if (parts.length === 3) {
        fechaFila = new Date(parts[2] + '-' + parts[1] + '-' + parts[0] + 'T00:00:00');
      }
    }

    // Aplicar filtros de fecha
    if (fInicio && fechaFila && fechaFila < fInicio) continue;
    if (fFin && fechaFila && fechaFila > fFin) continue;

    var turnosDetalle = fila[3] || "Sin registro";

    registros.push({
      cedula: fila[0],
      nombre: fila[1] || "Sin Nombre",
      centro: fila[2] || "Sin Centro",
      fecha: fila[4],
      turnosDetalle: turnosDetalle,
      horaInicio: fila[5],
      horaSalida: fila[6],
      tipoDiaInicio: fila[7],
      tipoDiaFin: fila[8],
      horasTotal: parseFloat(fila[9] || 0),
      hDiurNorm: parseFloat(fila[10] || 0),
      hNocNorm: parseFloat(fila[11] || 0),
      hDiurFest: parseFloat(fila[12] || 0),
      hNocFest: parseFloat(fila[13] || 0)
    });
  }

  return { status: 'ok', registros: registros };
}

/**
 * @summary Exporta la asistencia a CSV (delimitado por punto y coma).
 * @description Genera un string CSV listo para descargar.
 *
 * @param {Object} filtros - Filtros de fecha.
 * @returns {Object} { status: 'ok', filename: String, csvContent: String }.
 */
function exportarAsistenciaCSV(filtros) {
  var dataObj = obtenerDataAsistencia(filtros);
  if (dataObj.status !== 'ok') return dataObj;

  var registros = dataObj.registros;

  // Encabezado CSV
  var csv = "Cédula;Nombre;Centro;Fecha;Hora Inicio;Hora Salida;Rango Completo;Tipo Día Inicio;Tipo Día Fin;Total Horas;H.Ord.Diurna;H.Ord.Nocturna;H.Fest.Diurna;H.Fest.Nocturna\n";

  // Función auxiliar para formatear números (reemplazar punto por coma)
  var fmtNum = function(n) {
    return String(n).includes('.') ? String(n).replace('.', ',') : String(n);
  };

  // Función auxiliar para escapar comillas en CSV
  var escape = function(v) {
    if (v == null) return '""';
    var s = String(v);
    return '"' + s.replace(/"/g, '""') + '"';
  };

  for (var i = 0; i < registros.length; i++) {
    var r = registros[i];
    csv += [
      escape(r.cedula),
      escape(r.nombre),
      escape(r.centro),
      escape(r.fecha),
      escape(r.horaInicio),
      escape(r.horaSalida),
      escape(r.turnosDetalle),
      escape(r.tipoDiaInicio),
      escape(r.tipoDiaFin),
      fmtNum(r.horasTotal),
      fmtNum(r.hDiurNorm),
      fmtNum(r.hNocNorm),
      fmtNum(r.hDiurFest),
      fmtNum(r.hNocFest)
    ].join(';') + "\n";
  }

  var fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");

  return {
    status: 'ok',
    csvContent: csv,
    filename: 'Reporte_Asistencia_' + fechaActual + '.csv'
  };
}

// ===================================================================
// 14. FUNCIONES DE PRUEBA Y DIAGNÓSTICO
// ===================================================================

/**
 * @summary Prueba el cálculo de festivos para un año específico.
 * @description Ejecutar desde el editor para verificar festivos.
 */
function testFestivos2026() {
  var festivos2026 = _generarFestivos(2026);

  Logger.log('=== FESTIVOS COLOMBIA 2026 ===');
  Logger.log('Total: ' + festivos2026.size);

  // Convertir Set a Array y ordenar
  var lista = Array.from(festivos2026).sort();
  for (var i = 0; i < lista.length; i++) {
    Logger.log('  ' + lista[i]);
  }

  // Probar fechas específicas
  var cache = {};
  cache[2026] = festivos2026;

  Logger.log('=== PRUEBAS DE FECHAS ===');
  Logger.log('01/01/2026: ' + _esFestivo(new Date(2026, 0, 1), cache));   // Festivo (Año Nuevo)
  Logger.log('02/01/2026: ' + _esFestivo(new Date(2026, 0, 2), cache));   // Normal (Viernes)
  Logger.log('04/01/2026: ' + _esFestivo(new Date(2026, 0, 4), cache));   // Domingo
  Logger.log('05/01/2026: ' + _esFestivo(new Date(2026, 0, 5), cache));   // Normal (Lunes)
  Logger.log('12/01/2026: ' + _esFestivo(new Date(2026, 0, 12), cache));  // Festivo (Reyes - movido al lunes)
  Logger.log('25/12/2026: ' + _esFestivo(new Date(2026, 11, 25), cache)); // Festivo (Navidad)
}

/**
 * @summary Prueba el procesamiento de un turno de ejemplo.
 * @description Ejecutar desde el editor para verificar cálculo de horas.
 */
function testCalculoHoras() {
  var cache = {};

  // Turno de ejemplo: 02/01/2026 07:00 a 17:00 (10 horas, día normal)
  var turno = {
    inicio: new Date(2026, 0, 2, 7, 0, 0),  // 02/01/2026 07:00
    fin: new Date(2026, 0, 2, 17, 0, 0)     // 02/01/2026 17:00
  };

  var resultado = _calcularHorasPorTipo([turno], cache);

  Logger.log('=== TEST CÁLCULO HORAS ===');
  Logger.log('Turno: 02/01/2026 07:00 - 17:00');
  Logger.log('Tipo día: ' + _esFestivo(turno.inicio, cache));
  Logger.log('Total: ' + resultado.total);
  Logger.log('Diurnas Normales: ' + resultado.normalesDia);
  Logger.log('Nocturnas Normales: ' + resultado.normalesNoc);
  Logger.log('Diurnas Festivo: ' + resultado.festivosDia);
  Logger.log('Nocturnas Festivo: ' + resultado.festivosNoc);

  // Turno nocturno: 02/01/2026 22:00 a 03/01/2026 06:00 (8 horas, cruza medianoche)
  var turnoNoc = {
    inicio: new Date(2026, 0, 2, 22, 0, 0),  // 02/01/2026 22:00
    fin: new Date(2026, 0, 3, 6, 0, 0)       // 03/01/2026 06:00
  };

  var resultadoNoc = _calcularHorasPorTipo([turnoNoc], cache);

  Logger.log('');
  Logger.log('Turno: 02/01/2026 22:00 - 03/01/2026 06:00');
  Logger.log('Total: ' + resultadoNoc.total);
  Logger.log('Diurnas Normales: ' + resultadoNoc.normalesDia);
  Logger.log('Nocturnas Normales: ' + resultadoNoc.normalesNoc);
}
