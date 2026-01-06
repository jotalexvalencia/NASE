// ======================================================================
// 🗺️ mapaHerramientas.gs – Sistema de Menú, Mapas y Formato (NASE 2026)
// ----------------------------------------------------------------------
/**
 * @summary Módulo de interfaz de usuario (UI) y Visualización para Admin.
 * @description Gestiona el menú principal de Google Sheets, la visualización de mapas
 *              interactivos (Leaflet) y la aplicación de formato estético a las hojas.
 *
 * @features
 *   - 🎛️ Menú personalizado en la barra de herramientas ("NASE - Sistema").
 *   - 🗺️ Visualización de mapas comparativos (Empleado vs Centro) con radios de distancia.
 *   - 📏 Generación automática de tablas con formato "Zebra" (filas alternas).
 *   - 🔧 Integración con módulos externos (Geocodificación, Configuración).
 *
 * @author NASE Team
 * @version 1.6 (Versión unificada con Menú Admin optimizado)
 */

// ======================================================================
// 1. CONFIGURACIÓN DEL MENÚ PRINCIPAL
// ======================================================================

/**
 * @summary Crea el menú personalizado al abrir la hoja.
 * @description Esta función se ejecuta automáticamente (`onOpen` trigger) cuando se abre la hoja.
 *              Estructura los items en secciones lógicas para el usuario administrador.
 *
 * @param {Event} e - Objeto de evento de Apps Script (proporcionado por Sheets).
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🧭 NASE - Sistema');

  // ---------------------------------------------------------
  // 1. Geocodificación y Mapas
  // ---------------------------------------------------------
  // Geocodifica la fila seleccionada (usando geocodificador.gs)
  menu.addItem('➡ Geocodificar fila actual', 'geocodificarFilaActiva'); 
  menu.addSeparator();
  
  // Visualización de ubicaciones
  menu.addItem('🗺️ Ver mapa del registro', 'mostrarMapaDelRegistro');
  menu.addItem('📍 Ver mapa comparativo (centro vs registro)', 'mostrarMapaComparativo');
  menu.addSeparator();

  // ---------------------------------------------------------
  // 2. Utilidades Generales
  // ---------------------------------------------------------
  // ✅ Opción Visual: Lista de Salidas Pendientes (Sidebar)
  // Muestra la barra lateral con tabla y scroll (integración con SalidasPendientes.gs)
  menu.addItem('📋 Lista Salidas Pendientes (Visual)', 'mostrarListadoSalidasPendientes');

  // Nota: La opción "Filtrar registros" se ha movido al menú principal o se ejecuta como acción manual si se requiere.
  menu.addSeparator();

  // ---------------------------------------------------------
  // 3. Gestión de Asistencia y Registros
  // ---------------------------------------------------------
  // Genera la hoja de asistencia calculada (integración con asistencia.gs)
  menu.addItem('📊 Generar Asistencia', 'generarTablaAsistenciaSinValores');
  menu.addSeparator();

  // ---------------------------------------------------------
  // 4. Gestión de Salidas (Admin - Nuevas Opciones)
  // ---------------------------------------------------------
  // ✅ Opción A: Lista Visual (Sidebar)
  // Muestra la barra lateral con tabla y scroll.
  // menu.addItem('📋 Lista Salidas Pendientes (Visual)', 'mostrarListadoSalidasPendientes');

  menu.addSeparator();

  // ---------------------------------------------------------
  // 5. Configuración
  // ---------------------------------------------------------
  // Abre el panel lateral para configurar recargos nocturnos (config_horarios.html)
  menu.addItem('⚙️ Configurar Horas (Inicio y Fin del Recargo Nocturno)', 'mostrarConfiguracionHorarios');
  menu.addSeparator();

  // ---------------------------------------------------------
  // 6. Mantenimiento y Administración
  // ---------------------------------------------------------
  // Limpieza de triggers desactivados o antiguos
  menu.addItem('🧹 Limpieza Profunda de Triggers', 'limpiezaProfundaTriggers');
  
  // Opciones de mantenimiento (Comentadas por seguridad, descomentar si es necesario)
  // menu.addItem('🚀 Inicializar Sistema Optimizado', 'inicializarSistemaOptimizado');
  // menu.addItem('🗂️ Crear Archivo del Mes', 'createMonthlyArchive');

  menu.addToUi();
}

// ======================================================================
// 2. UTILIDADES INTERNAS (Helpers Robustos)
// ======================================================================

/**
 * @summary Busca el índice de una columna dinámicamente por nombre.
 * @description Es útil porque las columnas en Sheets pueden cambiar de orden.
 *              Busca ignorando mayúsculas/minúsculas y espacios.
 *
 * @param {Array} headers - Array de encabezados de la hoja (fila 1).
 * @param {Array} names - Array de nombres probables para buscar (ej: ['lat', 'latitud']).
 * @returns {Number} Índice de la columna (1-based), o -1 si no se encuentra.
 */
function _findHeaderIndex(headers, names) {
  if (!headers || !headers.length) return -1;
  const lower = headers.map(h => (h || "").toString().trim().toLowerCase());
  for (const cand of names) {
    const idx = lower.indexOf(cand.toString().trim().toLowerCase());
    if (idx !== -1) return idx + 1; // +1 porque Spreadsheet usa índices 1-based, Array usa 0-based
  }
  return -1;
}

/**
 * @summary Parsea (convierte) coordenadas de texto/número a Float.
 * @description Maneja valores extraños como fechas (1899...) y strings.
 *              Si el valor es "NO", "N/A" o similar, devuelve NaN.
 *
 * @param {String|Number} v - Valor de la celda (Latitud o Longitud).
 * @returns {Number} Coordenada decimal (Ej: 4.825) o NaN si es inválido.
 */
function _parseCoord(v) {
  if (v === null || typeof v === 'undefined') return NaN;
  let s = String(v).trim();
  
  // Limpieza básica: eliminar "[REVISAR]", espacios
  if (["[REVISAR]", "NO", "N/A"].includes(s.toUpperCase())) return NaN;
  
  // Reemplazar comas por puntos (formato Colombia)
  s = s.replace(",", ".");
  
  // Eliminar caracteres no numéricos excepto punto y signo
  s = s.replace(/[^0-9.\-]/g, "");
  
  let n = parseFloat(s);
  if (isNaN(n)) return NaN;
  
  // Corrección automática para coordenadas que se guardaron como fecha erróneamente (Ej: 1899...)
  // Si el número es muy grande (una fecha Excel es un número > 40000), dividir por 1^x para normalizar
  while (Math.abs(n) > 180) n /= 10;
  
  return n;
}

// ======================================================================
// 3. VISUALIZACIÓN DE MAPAS (Leaflet)
// ======================================================================

/**
 * @summary Muestra un mapa interactivo con el registro de la fila seleccionada.
 * @description Genera una ventana Modal con un mapa Leaflet.js.
 *              - Coloca un marcador rojo en la posición del empleado.
 *              - Busca el centro asignado en la hoja 'Centros'.
 *              - Coloca un marcador en la posición del centro (si existe).
 *              - Dibuja un círculo azul representando el radio de asistencia (Ej: 30m).
 *              - Muestra capas de "Callejero" y "Satélite".
 *
 * @requires 'Centros' sheet para obtener referencias (lat/lng ref, radio, URL imagen).
 */
function mostrarMapaDelRegistro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRes = ss.getActiveSheet();
  const fila = hojaRes.getActiveCell().getRow();

  // Validar contexto: No ejecutar en fila 1 (encabezados)
  if (!fila || fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una fila válida en la hoja Respuestas antes de usar 'Ver mapa del registro'.");
    return;
  }

  // ----------------------------------------------------------------------
  // 1. LEER DATOS DE LA FILA ACTIVA
  // ----------------------------------------------------------------------
  const headers = hojaRes.getRange(1, 1, 1, hojaRes.getLastColumn()).getValues()[0] || [];
  
  // Obtener índices de columnas dinámicamente
  const idxCentro = _findHeaderIndex(headers, ["centro"]);
  const idxCiudadCentro = _findHeaderIndex(headers, ["ciudad"]);
  const idxLat = _findHeaderIndex(headers, ["lat"]);
  const idxLng = _findHeaderIndex(headers, ["lng"]);
  const idxDir = _findHeaderIndex(headers, ["barrio / dirección", "direccion", "dirección"]);

  const filaVals = hojaRes.getRange(fila, 1, 1, hojaRes.getLastColumn()).getValues()[0];
  
  const centro = (filaVals[idxCentro - 1] || "").toString();
  const ciudadCentro = (filaVals[idxCiudadCentro - 1] || "").toString();
  
  // Parsear coordenadas usando la función auxiliar robusta
  const latEmp = _parseCoord(filaVals[idxLat - 1]);
  const lngEmp = _parseCoord(filaVals[idxLng - 1]);
  const direccion = (filaVals[idxDir - 1] || "").toString();

  // Validar que tengamos coordenadas válidas para pintar el mapa
  if (isNaN(latEmp) || isNaN(lngEmp)) {
    SpreadsheetApp.getUi().alert("⚠️ Coordenadas inválidas en la fila seleccionada. Verifica Lat/Lng en la hoja Respuestas.");
    return;
  }

  // ----------------------------------------------------------------------
  // 2. BUSCAR EL CENTRO EN LA HOJA 'Centros' (REFERENCIA)
  // ----------------------------------------------------------------------
  const hojaCentros = ss.getSheetByName("Centros");
  let latCentro = null, lngCentro = null, radio = 30, urlImagenCentro = "";
  
  if (hojaCentros) {
    const dataC = hojaCentros.getDataRange().getValues();
    const headersC = dataC[0].map(h => (h || "").toString().trim().toUpperCase());
    
    // Índices en la hoja Centros
    const idxC_Centro = _findHeaderIndex(headersC, ["centro"]) - 1;
    const idxC_Ciudad = _findHeaderIndex(headersC, ["ciudad"]) - 1;
    const idxC_Lat = _findHeaderIndex(headersC, ["lat ref", "latitud"]) - 1;
    const idxC_Lng = _findHeaderIndex(headersC, ["lng ref", "longitud"]) - 1;
    const idxC_Radio = _findHeaderIndex(headersC, ["radio"]) - 1;
    const idxC_UrlImagen = _findHeaderIndex(headersC, ["link_imagen", "url imagen", "imagen"]) - 1;

    // Iterar para encontrar el centro que coincide en nombre y ciudad
    for (let i = 1; i < dataC.length; i++) {
      const rowC = dataC[i];
      const nombreC = (rowC[idxC_Centro] || "").toString().trim();
      const ciudadC = (rowC[idxC_Ciudad] || "").toString().trim();
      
      if (nombreC.toUpperCase() === centro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
        latCentro = _parseCoord(rowC[idxC_Lat]);
        lngCentro = _parseCoord(rowC[idxC_Lng]);
        radio = rowC[idxC_Radio] ? Number(rowC[idxC_Radio]) : 30;
        urlImagenCentro = (idxC_UrlImagen >= 0 ? (rowC[idxC_UrlImagen] || "").toString().trim() : "");
        break;
      }
    }
  }

  // ----------------------------------------------------------------------
  // 3. GENERAR HTML DEL MAPA (Leaflet.js)
  // ----------------------------------------------------------------------
  // Construimos el código HTML y JavaScript necesario para el mapa
  const html = `
  <!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <!-- Leaflet CSS (Librería de mapas open source) -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
      body {margin:0;padding:10px;font-family:Arial,sans-serif}
      #map {height:680px;width:100%;border-radius:8px}
      .popup-img {width:100%;max-width:300px;border-radius:5px;margin-top:5px}
      .legend {background:#fff;padding:6px;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.15)}
    </style>
  </head>
  <body>
    <div id="map"></div>
    
    <!-- Leaflet JS -->
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    
    <script>
      // 1. Inicializar mapa centrado en la posición del empleado
      const map = L.map('map').setView([${latEmp}, ${lngEmp}], 17);
      
      // 2. Añadir capa de calles (OpenStreetMap) - Fondo por defecto
      const calle = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap'}).addTo(map);
      
      // 3. Añadir capa satelital (Esri World Imagery) - Como capa extra
      const sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{attribution:'© Esri'});
      
      // Control de capas (Switcher)
      L.control.layers({"🗺️ Callejero":calle,"🌍 Satélite":sat}).addTo(map);

      // 4. Crear Icono Personalizado (Marcador NASE)
      const iconoRegistro = L.divIcon({
        html:\`
          <svg height="36" width="36"><circle cx="18" cy="18" r="16" fill="#1976d2" stroke="white" stroke-width="2"/>
          <image href="https://raw.githubusercontent.com/jotalexvalencia/NASE/main/nase_marcador.png" x="10" y="10" width="16" height="16"/></svg>\`,
        iconSize:[36,36],iconAnchor:[18,36],popupAnchor:[0,-36] // Flecha apunta abajo
      });
      
      // 5. Añadir marcador del empleado
      const empMarker=L.marker([${latEmp},${lngEmp}],{icon:iconoRegistro}).addTo(map);
      empMarker.bindPopup("<b>📍 Registro Empleado</b><br>${centro}<br>${direccion}<br><b>Lat:</b> ${latEmp.toFixed(6)}<br><b>Lng:</b> ${lngEmp.toFixed(6)}").openPopup();

      // 6. Si se encontró el centro, añadir marcador y radio de seguridad
      if(${latCentro} && ${lngCentro}){
        const centroMarker=L.marker([${latCentro},${lngCentro}]).addTo(map);
        
        // Círculo representando el radio (30m por defecto)
        L.circle([${latCentro},${lngCentro}],{color:'#1e88e5',fillColor:'#1e88e5',fillOpacity:0.12,radius:${radio}}).addTo(map);
        
        // Popup con datos del centro y posible imagen
        let html="<b>🏢 Centro:</b> ${centro}<br>Radio: ${radio} m";
        if("${urlImagenCentro}") html+="<br><img src='${urlImagenCentro}' class='popup-img'>";
        
        centroMarker.bindPopup(html);
      }

      // 7. Control de escala
      L.control.scale().addTo(map);
    </script>
  </body></html>`;

  // Mostrar el HTML en un modal
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(880).setHeight(760),
    "🌍 Mapa del Registro"
  );
}

// ======================================================================
// 4. FORMATO ESTÁNDAR (Estilo Zebra)
// ======================================================================

/**
 * @summary Aplica formato profesional "Zebra" a una hoja.
 * @description Formatea encabezados (azul oscuro, texto blanco), alinea al centro,
 *              colorea filas pares e impares de forma distinta (Zebra), ajusta ancho de columnas
 *              y pone bordes a toda la tabla.
 * 
 * @param {Sheet} hoja - La hoja de Google Sheets a formatear.
 */
function aplicarFormatoEstandar(hoja) {
  if (!hoja) return;
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 1 || ultimaColumna < 1) return;

  // 1. Formato del Encabezado (Fila 1)
  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna);
  encabezados
    .setFontWeight("bold")
    .setBackground("#17365D")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // 2. Formato de las Filas de Datos
  for (let i = 2; i <= ultimaFila; i++) {
    hoja.getRange(i, 1, 1, ultimaColumna)
      // Color de fondo Zebra: Azul muy claro (2) o Blanco (1)
      .setBackground(i % 2 === 0 ? "#D9E1F2" : "#FFFFFF")
      // Alineación central y vertical
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  // 3. Bordes y Ajuste automático
  hoja.getDataRange().setBorder(true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  hoja.autoResizeColumns(1, ultimaColumna);
}

/**
 * @summary Función pública para formatear hojas clave.
 * @description Itera sobre una lista de hojas importantes ('Respuestas', 'Centros', 'Asistencia')
 *              y les aplica el formato estándar definido en `aplicarFormatoEstandar`.
 */
function formatearHojasEstandar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Lista de hojas que deben verse profesionales
  const hojas = ['Respuestas', 'Centros', 'Asistencia_SinValores'];
  
  hojas.forEach(n => {
    const hoja = ss.getSheetByName(n);
    if (hoja) aplicarFormatoEstandar(hoja);
  });
  
  SpreadsheetApp.getUi().alert("✅ Formato aplicado a todas las hojas estándar.");
}

// ======================================================================
// 5. CONFIGURACIÓN DE HORARIOS
// ======================================================================

/**
 * @summary Abre la plantilla de configuración de horarios en un Sidebar.
 * @description Carga el archivo HTML 'config_horarios' y lo muestra en una barra lateral
 *              para permitir al administrador configurar recargos salariales nocturnos.
 * 
 * @requires 'config_horarios.html' (Plantilla con Inputs de Hora Inicio/Fin).
 */
function mostrarConfiguracionHorarios() {
  // 1. Creamos la plantilla (Template)
  const template = HtmlService.createTemplateFromFile('config_horarios');
  
  // 2. Evaluamos la plantilla (Convierte de Template a Output)
  // ⚠️ IMPORTANTE: Solo después de .evaluate() existen los métodos .setWidth()
  const html = template.evaluate()
    .setWidth(700)  // ✅ Aquí ya no da error "is not a function"
    .setHeight(600);
  
  // 3. Mostramos el Sidebar
  // ✅ Pasamos solo el HtmlOutput. No pasamos el título como segundo argumento para evitar el error de parámetros.
  SpreadsheetApp.getUi().showSidebar(html);
}
