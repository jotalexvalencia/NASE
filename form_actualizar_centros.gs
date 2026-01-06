// ===================================================================
// 📁 form_actualizar_centros.gs – Backend Actualización de Centros (NASE 2026)
// -------------------------------------------------------------------
/**
 * @summary Módulo Backend para el formulario de actualización de GPS.
 * @description Gestiona el servidor para la página `actualizar_centros.html`.
 *              Valida la identidad del supervisor, carga la lista de centros pendientes
 *              y recibe las coordenadas GPS para actualizar la hoja 'Centros'.
 *
 * @features
 *   - 🎛️ **Inyección de Plantilla:** Sirve el HTML y le inyecta el objeto JSON
 *     de centros disponibles (`centrosInyectados`).
 *   - 🔐 **Validación de Supervisor:** Consulta la "BASE OPERATIVA" para verificar
 *     que el usuario sea un Supervisor activo.
 *   - 📋 **Filtrado de Pendientes:** Lee la hoja 'Centros' y genera un mapa de
 *     centros que NO tienen marcada la columna 'ACTUALIZADO' como 'Sí'.
 *   - ✍️ **Actualización Segura:** Usa `LockService` para evitar escrituras simultáneas.
 *
 * ✅ CORRECCIÓN v2.1: Búsqueda flexible de columnas LAT REF / LNG REF
 *
 * @author NASE Team
 * @version 2.1 (Columnas Flexibles)
 */

// ===================================================================
// 1. FUNCIÓN DE ENTRADA (SERVIDOR DE WEBAPP)
// ===================================================================

/**
 * @summary Sirve el formulario de actualización de centros.
 * @description Función `doGet()` principal para esta URL.
 *              Carga la plantilla HTML, obtiene la lista de centros pendientes,
 *              inyecta los datos en formato JSON y muestra el formulario al usuario.
 * 
 * @param {Event} e - Objeto de evento Apps Script.
 * @returns {HtmlOutput} La plantilla renderizada con los datos inyectados.
 */
function doGetActualizarCentrosPublico(e) {
  var template = HtmlService.createTemplateFromFile('actualizar_centros');
 
  // 1. Obtener datos de los centros pendientes de actualizar
  var datosCentros = obtenerCentrosPendientesPorActualizar();
  
  // 2. Validación defensiva
  if (!datosCentros) {
    datosCentros = {};
  }
 
  // 3. Serialización segura
  var jsonString = JSON.stringify(datosCentros);
 
  // 4. Asignamos a la variable del template
  template.centrosInyectados = jsonString;
 
  // 5. Retornamos la evaluación
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setTitle('NASE - Actualizar Ubicación');
}

// ===================================================================
// 2. VALIDACIÓN DE SUPERVISOR
// ===================================================================

/**
 * @summary Valida si una cédula pertenece a un supervisor activo.
 * @description Busca al usuario en la "BASE OPERATIVA".
 *              Verifica que su cargo sea 'SUPERVISOR' y su estado sea 'SUP'.
 * 
 * @param {String} cedula - Número de cédula del usuario.
 * @returns {Object} - { ok: boolean, nombre: String } o { ok: false, mensaje: String }.
 */
function buscarSupervisorPublico(cedula) {
  if (!cedula) {
    return { ok: false, mensaje: 'Cédula vacía' };
  }
 
  var cedulaLimpia = String(cedula).replace(/\D/g, '').trim();
  var supervisores = obtenerSupervisoresActivos();
  
  var encontrado = null;
  for (var i = 0; i < supervisores.length; i++) {
    if (supervisores[i].cedula === cedulaLimpia) {
      encontrado = supervisores[i];
      break;
    }
  }

  if (encontrado) {
    return { ok: true, nombre: encontrado.nombre };
  } else {
    return { ok: false, mensaje: 'Cédula no autorizada o inactiva.' };
  }
}

// ===================================================================
// 3. OBTENER DATOS DE CENTROS (PENDIENTES)
// ===================================================================

/**
 * @summary Obtiene la lista de centros disponibles para actualizar.
 * @description Lee la hoja 'Centros' y genera un mapa JSON.
 *              Omite los registros donde la columna 'ACTUALIZADO' contiene 'Sí'.
 * 
 * @returns {Object} Mapa JSON con clave "Ciudad|Centro" y valores { ciudad, centro, fila }.
 */
function obtenerCentrosPendientesPorActualizar() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = ss.getSheetByName(SHEET_CENTROS);
    
    if (!hoja) return {};
   
    var data = hoja.getDataRange().getValues();
    if (data.length < 2) return {}; 

    // 1. Obtener encabezados normalizados
    var headers = data[0].map(function(h) { return String(h).toUpperCase().trim(); });
   
    // 2. Localizar índices de columnas clave
    var idxCiudad = headers.findIndex(function(x) { return x.indexOf('CIUDAD') > -1; });
    if (idxCiudad === -1) idxCiudad = 0;

    var idxCentro = headers.findIndex(function(x) { 
      return x === 'CENTRO' || x.indexOf('CENTRO DE TRABAJO') > -1 || x.indexOf('SEDE') > -1; 
    });
    if (idxCentro === -1) idxCentro = 1;
   
    // Busca la columna "ACTUALIZADO"
    var idxActualizado = headers.findIndex(function(x) { 
      return x.indexOf('ACTUALIZADO') > -1; 
    });

    var centros = {};
   
    // 3. Recorrer filas de datos
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
     
      // Filtrar si ya está actualizado
      if (idxActualizado > -1) {
        var valorActualizado = String(row[idxActualizado] || '').trim().toLowerCase();
        if (valorActualizado === 'sí' || valorActualizado === 'si' || valorActualizado === 'yes') {
          continue;
        }
      }

      var ciudad = String(row[idxCiudad]).trim();
      var centroNombre = String(row[idxCentro]).trim();

      if (ciudad !== '' && centroNombre !== '') {
        var key = ciudad + '|' + centroNombre;
        centros[key] = {
          ciudad: ciudad,
          centro: centroNombre,
          fila: i + 1
        };
      }
    }
    
    return centros;
    
  } catch (e) {
    Logger.log("Error obteniendo centros: " + e.toString());
    return {};
  }
}

// ===================================================================
// 4. OBTENER LISTA DE SUPERVISORES
// ===================================================================

/**
 * @summary Obtiene la lista de supervisores activos desde RRHH.
 * @description Lee la "BASE OPERATIVA" y filtra por Cargo y Estado.
 *              Solo incluye usuarios con Cargo='...SUPERVISOR...' y Estado='SUP'.
 * 
 * @returns {Array<Object>} Array de objetos con `cedula` y `nombre`.
 */
function obtenerSupervisoresActivos() {
  try {
    var ssBase = SpreadsheetApp.openById(ID_LIBRO_BASE);
    var hoja = ssBase.getSheetByName('BASE OPERATIVA');
    
    if (!hoja) return [];
   
    var data = hoja.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = [];
    for (var h = 0; h < data[0].length; h++) {
      headers.push(String(data[0][h]).toUpperCase().trim());
    }

    var idxCedula = headers.findIndex(function(x) { return x.indexOf('DOCUMENTO') > -1; });
    var idxNombre = headers.findIndex(function(x) { return x.indexOf('NOMBRE') > -1; });
    var idxCargo = headers.findIndex(function(x) { return x.indexOf('CARGO') > -1; });
    var idxEstado = headers.findIndex(function(x) { return x.indexOf('ESTADO') > -1; });

    var supervisores = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var cargo = String(row[idxCargo] || '').toUpperCase();
      var estado = idxEstado > -1 ? String(row[idxEstado] || '').toUpperCase() : 'A';
     
      if (cargo.indexOf('SUPERVISOR') > -1 && (estado === 'SUP')) {
        supervisores.push({
          cedula: String(row[idxCedula] || '').replace(/\D/g, ''),
          nombre: String(row[idxNombre] || '').trim()
        });
      }
    }
    
    return supervisores;
  } catch (e) {
    Logger.log("Error obteniendo supervisores: " + e.toString());
    return [];
  }
}

// ===================================================================
// 5. ACTUALIZACIÓN DE COORDENADAS (CORREGIDA)
// ===================================================================

/**
 * @summary Escribe las nuevas coordenadas en la hoja Centros.
 * @description Función invocada por el botón "ACTUALIZAR" del HTML.
 *              - Valida que el centro esté pendiente.
 *              - Actualiza Latitud y Longitud.
 *              - Marca como "ACTUALIZADO".
 *              - Registra el nombre del supervisor que realizó la acción.
 * 
 * ✅ CORRECCIÓN: Búsqueda flexible de columnas (LAT REF, LNG REF, etc.)
 * 
 * @param {Object} dataInput - Objeto JSON del frontend `{ ciudad, centro, lat, lng, supervisor }`.
 * @returns {Object} - { status: 'ok/error', message: String }.
 */
function actualizarCoordenadasCentro(dataInput) {
  // 1. Bloqueo del Script (Mutex)
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { status: 'error', message: 'Servidor ocupado. Intente de nuevo.' };
  }

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = ss.getSheetByName(SHEET_CENTROS);
    
    if (!hoja) {
      return { status: 'error', message: 'Hoja Centros no encontrada.' };
    }
    
    // Re-obtener centros para validar
    var centrosPendientes = obtenerCentrosPendientesPorActualizar();
    
    // 2. Validar Clave (Ciudad + Centro)
    var key = dataInput.ciudad + '|' + dataInput.centro;
   
    if (!centrosPendientes[key]) {
      return { status: 'error', message: 'Centro ya actualizado o no encontrado.' };
    }

    var fila = centrosPendientes[key].fila;
    
    // Leer encabezados
    var headersRaw = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var headers = [];
    for (var k = 0; k < headersRaw.length; k++) {
      headers.push(String(headersRaw[k]).toUpperCase().trim());
    }
   
    // =====================================================================
    // ✅ CORRECCIÓN: BÚSQUEDA FLEXIBLE DE COLUMNAS
    // =====================================================================
    // Ahora busca cualquier columna que CONTENGA "LAT" o "LNG"
    // Esto funciona para: LAT, LATITUD, LAT REF, LATITUDE, etc.
    
    var idxLat = headers.findIndex(function(x) { 
      return x.indexOf('LAT') > -1; 
    });
    
    var idxLng = headers.findIndex(function(x) { 
      return x.indexOf('LNG') > -1 || x.indexOf('LON') > -1; 
    });
    
    var idxActualizado = headers.findIndex(function(x) { 
      return x.indexOf('ACTUALIZADO') > -1; 
    });
    
    var idxSupervisor = headers.findIndex(function(x) { 
      return x.indexOf('SUPERVISOR') > -1; 
    });

    // =====================================================================
    // DEBUG: Registrar columnas encontradas (útil para diagnóstico)
    // =====================================================================
    Logger.log('=== DEBUG COLUMNAS ===');
    Logger.log('Headers: ' + headers.join(' | '));
    Logger.log('idxLat: ' + idxLat + ' (columna: ' + (idxLat > -1 ? headers[idxLat] : 'NO ENCONTRADA') + ')');
    Logger.log('idxLng: ' + idxLng + ' (columna: ' + (idxLng > -1 ? headers[idxLng] : 'NO ENCONTRADA') + ')');
    Logger.log('idxActualizado: ' + idxActualizado);
    Logger.log('idxSupervisor: ' + idxSupervisor);
    Logger.log('Fila a actualizar: ' + fila);
    Logger.log('Datos recibidos: ' + JSON.stringify(dataInput));

    // Validar que existan columnas de coordenadas
    if (idxLat === -1) {
      return { 
        status: 'error', 
        message: 'No se encontró columna de Latitud. Columnas disponibles: ' + headers.join(', ')
      };
    }
    
    if (idxLng === -1) {
      return { 
        status: 'error', 
        message: 'No se encontró columna de Longitud. Columnas disponibles: ' + headers.join(', ')
      };
    }

    // 3. ESCRIBIR COORDENADAS
    // Convertir a número para guardar correctamente
    var latNum = parseFloat(String(dataInput.lat).replace(',', '.'));
    var lngNum = parseFloat(String(dataInput.lng).replace(',', '.'));
    
    if (isNaN(latNum) || isNaN(lngNum)) {
      return { status: 'error', message: 'Coordenadas inválidas.' };
    }
    
    // Escribir en las columnas correspondientes (idxLat + 1 porque getRange es 1-based)
    hoja.getRange(fila, idxLat + 1).setValue(latNum);
    hoja.getRange(fila, idxLng + 1).setValue(lngNum);
    
    // 4. MARCAR COMO ACTUALIZADO
    if (idxActualizado > -1) {
      hoja.getRange(fila, idxActualizado + 1).setValue('Sí');
    }
    
    // 5. REGISTRAR SUPERVISOR
    if (idxSupervisor > -1) {
      hoja.getRange(fila, idxSupervisor + 1).setValue(dataInput.supervisor);
    }
   
    // Forzar escritura inmediata
    SpreadsheetApp.flush();
    
    Logger.log('✅ Centro actualizado: ' + key + ' -> Lat: ' + latNum + ', Lng: ' + lngNum);
    
    return { status: 'ok', message: 'Centro actualizado correctamente' };

  } catch (e) {
    Logger.log('❌ Error actualizando centro: ' + e.toString());
    return { status: 'error', message: 'Error interno: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================
// 6. FUNCIÓN DE DIAGNÓSTICO
// ===================================================================

/**
 * @summary Diagnostica la estructura de la hoja Centros.
 * @description Ejecutar desde el editor para verificar que las columnas
 *              se detectan correctamente.
 */
function diagnosticarHojaCentros() {
  Logger.log('');
  Logger.log('╔══════════════════════════════════════════════════════════════╗');
  Logger.log('║        🔍 DIAGNÓSTICO HOJA CENTROS                           ║');
  Logger.log('╚══════════════════════════════════════════════════════════════╝');
  Logger.log('');
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = ss.getSheetByName(SHEET_CENTROS);
    
    if (!hoja) {
      Logger.log('❌ Hoja "' + SHEET_CENTROS + '" no encontrada');
      return;
    }
    
    var headersRaw = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var headers = [];
    for (var k = 0; k < headersRaw.length; k++) {
      headers.push(String(headersRaw[k]).toUpperCase().trim());
    }
    
    Logger.log('📋 Columnas encontradas (' + headers.length + '):');
    Logger.log('');
    
    for (var i = 0; i < headers.length; i++) {
      var letra = String.fromCharCode(65 + i); // A, B, C, D...
      Logger.log('   ' + letra + ': ' + headers[i]);
    }
    
    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🔎 BÚSQUEDA DE COLUMNAS CLAVE:');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    
    // Buscar columnas
    var idxCiudad = headers.findIndex(function(x) { return x.indexOf('CIUDAD') > -1; });
    var idxCentro = headers.findIndex(function(x) { return x === 'CENTRO' || x.indexOf('CENTRO') > -1; });
    var idxLat = headers.findIndex(function(x) { return x.indexOf('LAT') > -1; });
    var idxLng = headers.findIndex(function(x) { return x.indexOf('LNG') > -1 || x.indexOf('LON') > -1; });
    var idxRadio = headers.findIndex(function(x) { return x.indexOf('RADIO') > -1; });
    var idxActualizado = headers.findIndex(function(x) { return x.indexOf('ACTUALIZADO') > -1; });
    var idxSupervisor = headers.findIndex(function(x) { return x.indexOf('SUPERVISOR') > -1; });
    
    var columnas = [
      { nombre: 'CIUDAD', idx: idxCiudad },
      { nombre: 'CENTRO', idx: idxCentro },
      { nombre: 'LATITUD', idx: idxLat },
      { nombre: 'LONGITUD', idx: idxLng },
      { nombre: 'RADIO', idx: idxRadio },
      { nombre: 'ACTUALIZADO', idx: idxActualizado },
      { nombre: 'SUPERVISOR', idx: idxSupervisor }
    ];
    
    for (var j = 0; j < columnas.length; j++) {
      var col = columnas[j];
      var letra = col.idx > -1 ? String.fromCharCode(65 + col.idx) : '-';
      var header = col.idx > -1 ? headers[col.idx] : 'NO ENCONTRADA';
      var icono = col.idx > -1 ? '✅' : '❌';
      
      Logger.log('   ' + icono + ' ' + col.nombre + ': Columna ' + letra + ' (' + header + ')');
    }
    
    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📊 DATOS DE MUESTRA (primeras 3 filas):');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    
    var datos = hoja.getRange(2, 1, Math.min(3, hoja.getLastRow() - 1), hoja.getLastColumn()).getValues();
    
    for (var m = 0; m < datos.length; m++) {
      var ciudad = idxCiudad > -1 ? datos[m][idxCiudad] : '-';
      var centro = idxCentro > -1 ? datos[m][idxCentro] : '-';
      var lat = idxLat > -1 ? datos[m][idxLat] : '-';
      var lng = idxLng > -1 ? datos[m][idxLng] : '-';
      var actualizado = idxActualizado > -1 ? datos[m][idxActualizado] : '-';
      
      Logger.log('   Fila ' + (m + 2) + ': ' + ciudad + ' | ' + centro + ' | Lat: ' + lat + ' | Lng: ' + lng + ' | Act: ' + actualizado);
    }
    
    Logger.log('');
    Logger.log('✅ Diagnóstico completado');
    
  } catch (e) {
    Logger.log('❌ Error: ' + e.toString());
  }
}
