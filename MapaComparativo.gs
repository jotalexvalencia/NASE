// ======================================================================
// 📍 MapaComparativo.gs – Visualización Comparativa (NASE 2026)
// ----------------------------------------------------------------------
/**
 * @summary Módulo de visualización geoespacial comparativa.
 * @description Genera un mapa interactivo (usando Leaflet.js) que compara
 *              la ubicación de un empleado contra su centro de trabajo asignado.
 *
 * @features
 *   - 🗺️ Mapa Leaflet: Muestra capas de "Callejero" y "Satélite".
 *   - 🏢 Centro: Marcador del centro de trabajo con popup (Dirección e Imagen).
 *   - 📍 Empleado: Marcador de la ubicación del empleado (Lat/Lng).
 *   - 📏 Distancia: Cálculo de la distancia en metros (Fórmula Haversine).
 *   - 🎨 Círculo: Dibuja el radio de asistencia del centro (mínimo 80m visual).
 *   - 🚦 Semáforo: El marcador del empleado cambia de color:
 *       - 🔵 Azul si está DENTRO del radio (Distancia <= Radio).
 *       - 🔴 Rojo si está FUERA del radio (Distancia > Radio).
 *   - 🧭 Línea: Dibuja una línea punteada conectando el centro con el empleado.
 *   - 🧭 Leyenda: Muestra un panel de control con detalles de la distancia.
 *
 * @author NASE Team
 * @version 1.1 (Clean Code + Documentación)
 */

// ======================================================================
// 1. FUNCIÓN PRINCIPAL: Generar Mapa Comparativo
// ======================================================================

/**
 * @summary Genera y muestra el mapa comparativo en una ventana modal.
 * @description Lee la fila seleccionada, busca el centro correspondiente,
 *              calcula la distancia y construye el código HTML para Leaflet.
 * 
 * @workflow
 *   1. Lee datos de la fila activa (Cédula, Centro, Lat, Lng).
 *   2. Busca los datos del centro en la hoja 'Centros' (Lat Ref, Lng Ref, Radio, Imagen).
 *   3. Calcula la distancia en metros entre el empleado y el centro.
 *   4. Genera un string HTML conteniendo CSS, estructura DIV y código JS de Leaflet.
 *   5. Muestra el resultado en un Modal Dialog de Google Sheets.
 * 
 * @requires Hoja 'Centros' con columnas: Centro, Ciudad, Lat Ref, Lng Ref, Radio, Link Imagen.
 */
function mostrarMapaComparativo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResp = ss.getActiveSheet();
  const fila = hojaResp.getCurrentCell().getRow();

  // Validar contexto: Seleccionar una fila con datos (no fila 1)
  if (!fila || fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una fila válida en la hoja Respuestas antes de usar 'Ver mapa comparativo'.");
    return;
  }

  // ================================================================
  // 1️⃣ OBTENER DATOS DEL EMPLEADO (Fila Activa)
  // ================================================================
  const headersResp = hojaResp.getRange(1, 1, 1, hojaResp.getLastColumn()).getValues()[0] || [];
  
  // Función auxiliar local para encontrar índices de columnas
  const findIdx = (hdrs, candidates) => {
    const low = hdrs.map(h => (h || "").toString().trim().toLowerCase());
    for (const cand of candidates) {
      const idx = low.indexOf(cand.toString().trim().toLowerCase());
      if (idx !== -1) return idx + 1;
    }
    return -1;
  };

  // Leer datos de la fila
  const idxCentro = findIdx(headersResp, ["centro", "centros", "centro de trabajo"]);
  const idxCiudadCentro = findIdx(headersResp, ["ciudad", "city"]);
  const idxLat = findIdx(headersResp, ["lat", "latitude", "latitud"]);
  const idxLng = findIdx(headersResp, ["lng", "lon", "long", "longitude", "longitud"]);
  const idxDir = findIdx(headersResp, ["barrio / dirección", "direccion", "dirección", "address"]);

  const filaVals = hojaResp.getRange(fila, 1, 1, hojaResp.getLastColumn()).getValues()[0];
  
  const centro = (filaVals[idxCentro - 1] || "").toString().trim();
  const ciudadCentro = (filaVals[idxCiudadCentro - 1] || "").toString().trim();
  const latEmp = _parseCoord(filaVals[idxLat - 1]); // Helper global
  const lngEmp = _parseCoord(filaVals[idxLng - 1]); // Helper global
  const direccion = (filaVals[idxDir - 1] || "").toString();

  // Validar datos mínimos (Coordenadas y Centro)
  if (isNaN(latEmp) || isNaN(lngEmp)) {
    SpreadsheetApp.getUi().alert("⚠️ Coordenadas inválidas en la fila seleccionada. Verifica Lat/Lng en la hoja Respuestas.");
    return;
  }
  if (!centro) {
    SpreadsheetApp.getUi().alert("⚠️ La fila seleccionada no tiene un Centro asignado.");
    return;
  }

  // ================================================================
  // 2️⃣ BUSCAR INFORMACIÓN DEL CENTRO DE REFERENCIA (Hoja 'Centros')
  // ================================================================
  const hojaCentros = ss.getSheetByName("Centros");
  if (!hojaCentros) { SpreadsheetApp.getUi().alert("⚠️ La hoja 'Centros' no existe."); return; } 
  
  const dataCentros = hojaCentros.getDataRange().getValues();
  const headersCentros = dataCentros[0];
  
  // Función auxiliar para buscar en la hoja Centros
  const getColC = (name) => _findHeaderIndex(headersCentros, [name]);

  let latCentro = null, lngCentro = null, radio = 30, urlImagenCentro = "", direccionCentro = "";
  
  // Iterar para encontrar el centro que coincida con el del empleado (Nombre + Ciudad)
  for (let i = 1; i < dataCentros.length; i++) {
    const rowC = dataCentros[i];
    const nombreC = (rowC[getColC("centro") - 1] || "").toString().trim();
    const ciudadC = (rowC[getColC("ciudad") - 1] || "").toString().trim();
   
    // Coincidencia insensible a mayúsculas
    if (nombreC.toUpperCase() === centro.toUpperCase() && ciudadC.toUpperCase() === ciudadCentro.toUpperCase()) {
      latCentro = _parseCoord(rowC[getColC("lat ref") - 1]);
      lngCentro = _parseCoord(rowC[getColC("lng ref") - 1]);
      radio = rowC[getColC("radio") - 1] ? Number(rowC[getColC("radio") - 1]) : 30;
      direccionCentro = rowC[getColC("direccion") - 1] || rowC[getColC("barrio / dirección") - 1] || "";
      urlImagenCentro = (getColC("link_imagen") >= 0 ? (rowC[getColC("link_imagen") - 1] || "").toString().trim() : "");
      break;
    }
  }

  // Validar que se encontró el centro de referencia
  if (isNaN(latCentro) || isNaN(lngCentro))
    return SpreadsheetApp.getUi().alert(`⚠️ Coordenadas del centro "${centro}" no válidas o faltantes en la hoja 'Centros'.`);

  // ================================================================
  // 3️⃣ CÁLCULO DE DISTANCIA Y ESTADO (Dentro/Fuera)
  // ================================================================
  const distancia = _distMetros(latCentro, lngCentro, latEmp, lngEmp); // Helper global
  const dentro = distancia <= radio;

  // ================================================================
  // 4️⃣ GENERAR HTML DEL MAPA (Leaflet.js)
  // ================================================================
  const leafletCSS = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
  const leafletJS = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";

  /**
   * Construcción dinámica del HTML.
   * Se inyectan las variables JS (${latEmp}, etc.) directamente en el string de HTML.
   */
  const html = `
  <!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <!-- Hoja de estilos de Leaflet -->
    <link rel="stylesheet" href="${leafletCSS}" />
    <style>
      body {margin:0;padding:10px;font-family:Arial,sans-serif}
      #map {height:720px;width:100%;border-radius:8px}
      /* Estilos para popups de información */
      .popup-img {width:100%;max-width:300px;border-radius:5px;margin-top:5px}
      .legend {background:#fff;padding:6px;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,0.12)}
    </style>
  </head>
  <body>
    <div class="info">🏢 Centro: ${centro}</div>
    <div id="map"></div>
    
    <!-- Carga del script JS de Leaflet -->
    <script src="${leafletJS}"></script>
    <script>
      // Inicializar mapa: Centrado en el punto medio entre Centro y Empleado, Zoom 13
      const map = L.map('map').setView([ ((${latCentro}+${latEmp})/2), ((${lngCentro}+${lngEmp})/2) ], 13);
      
      // Capa 1: Callejero (OpenStreetMap)
      const calle = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap'}).addTo(map);
      
      // Capa 2: Satélite (Esri World Imagery)
      const sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{attribution:'© Esri'});
      
      // Control de capas (Switcher)
      L.control.layers({"🗺️ Callejero":calle,"🌍 Satélite":sat}).addTo(map);

      // -------------------------------------------------------------
      // MARCADOR DEL CENTRO DE TRABAJO
      // -------------------------------------------------------------
      const centroMarker = L.marker([${latCentro}, ${lngCentro}]).addTo(map);
      
      let htmlCentro = "<b>🏢 Centro de Trabajo</b><br>${centro}<br><b>Ciudad:</b> ${ciudadCentro}<br><b>Dirección:</b> ${direccionCentro}<br><b>Radio:</b> ${radio} m";
      // Agregar imagen si existe
      if ("${urlImagenCentro}") htmlCentro += "<br><img src='${urlImagenCentro}' class='popup-img'>";
      
      centroMarker.bindPopup(htmlCentro);

      // -------------------------------------------------------------
      // CÍRCULO DE RADIO (Visualización de Área Permitida)
      // -------------------------------------------------------------
      // Se usa Math.max(radio, 80) para que el círculo sea visible aunque el radio sea muy pequeño (<10m)
      L.circle([${latCentro}, ${lngCentro}],{
        color:'#2e7d32',fillColor:'#2e7d32',fillOpacity:0.12,
        radius:Math.max(${radio},80)
      }).addTo(map);

      // -------------------------------------------------------------
      // MARCADOR DEL EMPLEADO
      // -------------------------------------------------------------
      // Color dinámico según estado: Azul (#1976d2) si está dentro, Rojo (#d32f2f) si está fuera
      const color = ${dentro} ? '#1976d2' : '#d32f2f';
      
      const icono = L.divIcon({
        html: \`
          <svg height="36" width="36"><circle cx="18" cy="18" r="16" fill="\${color}" stroke="white" stroke-width="2"/></svg>
          <image href="https://raw.githubusercontent.com/jotalexvalencia/NASE/main/nase_marcador.png" x="10" y="10" width="16" height="16"/></svg>\`,
        iconSize:[36,36],iconAnchor:[18,36],popupAnchor:[0,-36]
      });
      
      const emp = L.marker([${latEmp}, ${lngEmp}], {icon:icono}).addTo(map);
      
      // Estado visual para el popup
      const estado = ${dentro} ? '✅ DENTRO del radio' : '❌ FUERA del radio';
      emp.bindPopup("<b>📍 Registro Empleado</b><br>${centro}<br>${direccion}<br><b>Lat:</b> ${latEmp.toFixed(6)}<br><b>Lng:</b> ${lngEmp.toFixed(6)}<br><b>Estado:</b> "+estado+"<br><b>Distancia:</b> ${distancia.toFixed(1)} m");

      // -------------------------------------------------------------
      // LÍNEA DE CONEXIÓN (Centro -> Empleado)
      // -------------------------------------------------------------
      L.polyline([[${latCentro}, ${lngCentro}],[${latEmp}, ${lngEmp}]],{color:'gray',weight:2,dashArray:'6,6'}).addTo(map);

      // -------------------------------------------------------------
      // LEYENDA PERSONALIZADA
      // -------------------------------------------------------------
      const legend = L.control({position:'bottomleft'});
      legend.onAdd = function (map) {
        const div = L.DomUtil.create('div', 'legend');
        div.innerHTML = "<b>📋 Leyenda</b><br>🏢 Centro<br>🔘 Radio: ${radio} m<br>🔵/🔴 Registro(dentro)/Registro(fuera)<br><b>Distancia:</b> ${distancia.toFixed(1)} m";
        return div;
      };
      legend.addTo(map);
    </script>
  </body>
  </html>`;

  // Mostrar el HTML en un Modal (Ventana emergente)
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(920).setHeight(780),
    "📍 Mapa Comparativo (Centro vs Registro)"
  );
}

// ======================================================================
// 2. UTILIDADES LOCALES / GLOBALES (Helpers)
// ======================================================================

/**
 * @summary Busca índice de encabezado (Versión Global).
 * @description Esta función está definida globalmente o se asume su existencia
 *              para evitar duplicación de código en el archivo.
 */
function _findHeaderIndex(headers, names) {
  if (!headers || !headers.length) return -1;
  const lower = headers.map(h => (h || "").toString().trim().toLowerCase());
  for (const cand of names) {
    const idx = lower.indexOf(cand.toString().trim().toLowerCase());
    if (idx !== -1) return idx + 1;
  }
  return -1;
}

/**
 * @summary Normaliza coordenadas (Versión Global).
 * @description Convierte strings, fechas o números a Float válido.
 *              Maneja errores comunes de formato en Sheets.
 */
function _parseCoord(v) {
  if (v == null || typeof v === 'undefined') return NaN;
  let s = String(v).trim();
  if (!s) return NaN;
  
  // Limpiar texto no numérico
  if (["[REVISAR]", "NO", "N/A"].includes(s.toUpperCase())) return NaN;
  
  s = s.replace(/\s+/g, "").replace(",", "."); // Coma a punto (formato Colombia)
  s = s.replace(/[^0-9.\-]/g, ""); // Eliminar caracteres extra excepto punto y menos
  
  let n = parseFloat(s);
  if (isNaN(n)) return NaN;
  
  // Corrección si el número es una fecha (ej: 1899-12-30 hora)
  while (Math.abs(n) > 180) n /= 10;
  
  return n;
}

/**
 * @summary Calcula la distancia Haversine en metros (Versión Global).
 * @description Fórmula trigonométrica para calcular la distancia más corta
 *              entre dos puntos en una esfera.
 * 
 * @param {Number} aLat, aLng - Coordenadas del punto A.
 * @param {Number} bLat, bLng - Coordenadas del punto B.
 * @returns {Number} Distancia en metros.
 */
function _distMetros(aLat, aLng, bLat, bLng) {
  const R = 6371000; // Radio de la Tierra en metros
  const dLat = (bLat - aLat) * Math.PI / 180;
  const dLng = (bLng - aLng) * Math.PI / 180;
  const A =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(aLat * Math.PI / 180) * Math.cos(bLat * Math.PI / 180) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  const c = 2 * R * Math.atan2(Math.sqrt(A), Math.sqrt(1 - A));
  return R * c;
}
