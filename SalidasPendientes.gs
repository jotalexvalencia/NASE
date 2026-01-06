// ======================================================================
// 📋 SalidasPendientes.gs – Visualización y Corrección de Turnos (NASE 2026)
// ----------------------------------------------------------------------
/**
 * @summary Módulo de Administración para cierre de turnos (Entrada/Salida).
 * @description Este es el módulo administrativo "Visual y Robusto". Centraliza la
 *              gestión de turnos abiertos en una interfaz Sidebar.
 * 
 * @features (Soluciones Q1, Q2, Q3)
 *   - 🎨 **Formato Limpio (Q1):** Convierte fechas tipo "Fri Jan 02..." a "dd/mm/yyyy".
 *   - ✍️ **Validación Manual (Q2):** Permite ingresar fecha/hora en formato "dd/mm/yyyy hh:mm".
 *     Validación robusta que evita errores de zona horaria (interpreta 02/01 como Ene 2).
 *   - 🎯 **Detección de Basura:** Identifica celdas con "Jan 02 00:00" (fecha por defecto)
 *     y las trata como "Pendientes" para permitir su corrección.
 *   - 📊 **Centralización (Q3):** La función del menú llama solo al Sidebar.
 *     Se elimina la opción de menú individual.
 *   - 👁️ **Auditoría:** Registra todos los cambios en una hoja dedicada.
 *
 * @author NASE Team
 * @version 2.0 (Visual + Robusta)
 */

// ======================================================================
// 1. CONFIGURACIÓN DEL SISTEMA
// ======================================================================

const SP_SHEET_RESPUESTAS = 'Respuestas';
const SP_TZ = "America/Bogota";

// Modo de Auditoría (Sheet = Guarda en hoja, None = No hace nada)
const SP_ADMIN_AUDIT_MODE = "sheet"; 
const SP_AUDIT_SHEET = "Auditoria_Salidas_Admin";

// -------------------------------------------------------------------
// 1.1 ENCABEZADOS OFICIALES HOJA RESPUESTAS (Mapeo de Índices)
// -------------------------------------------------------------------
const SP_HEADERS = [
  "Cédula",
  "Centro",
  "Ciudad",
  "Lat",
  "Lng",
  "Acepto",
  "Ciudad_Geo",
  "Dir_Geo",
  "Accuracy",
  "Dentro",      // Columna J -> DENTRO ENTRADA
  "Distancia",
  "Observaciones",
  "Nombre",
  "Foto",
  "Fecha Entrada", // 14
  "Hora Entrada",   // 15
  "Foto Entrada",
  "Fecha Salida", // 17
  "Hora Salida",   // 18
  "Foto Salida",
  "Dentro Salida"  // 20 (NUEVA COLUMNA)
];

/**
 * @summary Mapa de Índices (0-based) para acceso rápido a columnas.
 * @description Alineado con el array `RESP_HEADERS` del sistema principal.
 *              Se usa en este archivo para leer y escribir en 'Respuestas'.
 */
const SP_I = {
  CED: 0,
  CENTRO: 1,
  CIUDAD: 2,
  LAT: 3,
  LNG: 4,
  ACEPTO: 5,
  OBS: 11,
  NOMBRE: 12,
  FOTO: 13,
  FECHA_ENT: 14,
  HORA_ENT: 15,
  FOTO_ENT: 16,
  FECHA_SAL: 17,
  HORA_SAL: 18,
  FOTO_SAL: 19,
  DENTRO_SAL: 20
};

// ======================================================================
// 2. UTILIDADES DE FORMATEO (Visual Limpio - Q1)
// ======================================================================

/**
 * @summary Asegura que la hoja 'Respuestas' tenga los encabezados correctos.
 * @private
 */
function spAsegurarEncabezados_(hoja) {
  const lastCol = hoja.getLastColumn();
  if (lastCol === 0) {
    hoja.getRange(1, 1, 1, SP_HEADERS.length).setValues([SP_HEADERS]);
    return;
  }
  const current = hoja.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const currentLen = current.length;
  
  // Si faltan columnas al final, agregarlas
  if (currentLen < SP_HEADERS.length) {
    const faltantes = SP_HEADERS.slice(currentLen);
    hoja.getRange(1, currentLen + 1, 1, faltantes.length).setValues([faltantes]);
  }
}

/**
 * @summary Convierte una fecha al formato visual Limpio DD/MM/YYYY.
 * @description Soluciona el problema de fechas tipo "Fri Jan 02..." o "1899...".
 *              Si es Objeto Date, lo formatea. Si es String, lo limpia.
 * 
 * @param {Date|String} valor - Valor de la celda de fecha.
 * @returns {String} Fecha en formato dd/mm/yyyy.
 */
function spFormatearFecha_(valor) {
  if (!valor) return '';
  
  // Si es Objeto Date nativo (formateado en Sheets), convertimos a string limpio
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, SP_TZ, "dd/MM/yyyy");
  }
  
  // Si es string
  const str = String(valor || '').trim();
  
  // Si ya está en formato dd/mm/yyyy, devolverlo tal cual (Limpieza visual)
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
    return str;
  }
  
  // Intentar parsear como fecha y reformatear (para manejo de fechas extrañas)
  try {
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, SP_TZ, "dd/MM/yyyy");
    }
  } catch(e) {}
  
  return str; // Fallback si no se puede interpretar
}

/**
 * @summary Convierte una hora al formato visual Limpio HH:mm:ss.
 * @description Maneja objetos Date Time (que incluyen 1899) y strings.
 *              Extrae HH:mm:ss de forma segura.
 * 
 * @param {Date|String} valor - Valor de la celda de hora.
 * @returns {String} Hora en formato HH:mm:ss.
 */
function spFormatearHora_(valor) {
  if (!valor) return '';
  
  // Si es Objeto Date
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, SP_TZ, "HH:mm:ss");
  }
  
  // Si es string
  const str = String(valor || '').trim();
  
  // Si ya está en formato HH:mm:ss o HH:mm, devolverlo
  if (/^\d{2}:\d{2}(:\d{2})?$/.test(str)) {
    return str;
  }
  
  // Intentar parsear y extraer hora
  try {
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, SP_TZ, "HH:mm:ss");
    }
  } catch(e) {}
  
  return str;
}

/**
 * @summary Parsea un string "DD/MM/YYYY HH:MM" (o HH:MM:SS) a Objeto Date.
 * @description Soluciona Q2: Validación manual robusta.
 *              Convierte manualmente a formato ISO para evitar errores de zona horaria
 *              (Ej: 02/01/2026 se interpreta como 2 de Enero, no 1 de Febrero).
 * 
 * @param {String} textoFechaHora - Fecha y hora ingresada por el admin (espacio separado).
 * @returns {Date|null} Objeto Date o null si el formato es inválido.
 */
function spParsearFechaHora_(textoFechaHora) {
  if (!textoFechaHora) return null;
  
  const texto = String(textoFechaHora).trim();
  
  // Formato esperado: "DD/MM/YYYY HH:MM" o "DD/MM/YYYY HH:MM:SS"
  // El usuario ingresa con espacio entre fecha y hora.
  const partes = texto.split(' ');
  
  if (partes.length < 2) return null;
  
  const fechaParts = partes[0].split('/'); // [DD, MM, YYYY]
  const horaParts = partes[1].split(':');  // [HH, MM, SS]
  
  if (fechaParts.length !== 3 || horaParts.length < 2) return null;
  
  const dia = fechaParts[0].padStart(2, '0');
  const mes = fechaParts[1].padStart(2, '0');
  const anio = fechaParts[2];
  const hora = horaParts[0].padStart(2, '0');
  const min = horaParts[1].padStart(2, '0');
  const seg = horaParts[2] ? horaParts[2].padStart(2, '0') : '00';
  
  // Crear fecha en formato ISO (YYYY-MM-DDTHH:mm:ss) para evitar ambigüedades de zona horaria
  // 2026-01-02T14:30:00
  const isoStr = `${anio}-${mes}-${dia}T${hora}:${min}:${seg}`;
  const dt = new Date(isoStr);
  
  return isNaN(dt.getTime()) ? null : dt;
}

// ======================================================================
// 3. FUNCIÓN PRINCIPAL DEL MENÚ (Visual Limpia - Q3)
// ======================================================================

/**
 * @summary Genera y muestra el Sidebar de salidas pendientes.
 * @description
 *   - Lee los datos usando `getRawPendingList_` (con formato limpio aplicado).
 *   - Genera un HTML con una tabla de registros abiertos.
 *   - Muestra la fecha/hora en formato dd/mm/yyyy HH:mm:ss (Solución Q1).
 *   - Permite cerrar turnos directamente desde el Sidebar (Centralización Q3).
 * 
 * @see getRawPendingList_() para la lógica de datos.
 * @see corregirDesdeSidebar() para la lógica de cierre.
 */
function mostrarListadoSalidasPendientes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SP_SHEET_RESPUESTAS);

  if (!hoja) { ui.alert("❌ No existe la hoja 'Respuestas'."); return; }
  
  // Validar estructura de columnas
  spAsegurarEncabezados_(hoja);

  // Obtener lista de pendientes (con formato limpio)
  const data = getRawPendingList_(); 
  
  if (!data || data.length === 0) {
    ui.alert("✅ No hay salidas pendientes. Todos los turnos están cerrados.");
    return;
  }

  // -----------------------------------------------------------
  // CONSTRUCCIÓN HTML (Visual Limpia)
  // -----------------------------------------------------------
  let html = '<style>';
  html += 'body { font-family: Arial, sans-serif; font-size: 12px; } ';
  html += 'h3 { color: #17365D; font-size: 18px; margin-bottom: 10px; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px; } ';
  html += '.table-container { max-height: 500px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px; } ';
  html += 'table { width: 100%; border-collapse: collapse; } ';
  html += 'th { background-color: #f8f9fa; color: #17365D; padding: 10px; text-align: left; position: sticky; top: 0; border-bottom: 2px solid #eee; } ';
  html += 'td { padding: 10px; border-bottom: 1px solid #f1f3f5; vertical-align: middle; } ';
  html += 'tr:hover { background-color: #f1f8ff; } ';
  html += '.btn-cerrar { background-color: #ffc107; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-weight: bold; } ';
  html += '.btn-cerrar:hover { background-color: #e0a800; } ';
  html += '</style>';

  html += '<h3>📋 Salidas Pendientes (Admin)</h3>';
  html += '<p>Lista de turnos abiertos. Presiona "🔒 Cerrar" para cerrar turno.</p>';
  
  html += '<div class="table-container"><table>';
  html += '<thead><tr><th>Cédula</th><th>Nombre</th><th>Centro</th><th>Entrada (Fecha/Hora)</th><th>Acción</th></tr></thead>';
  html += '<tbody>';

  // Iterar sobre los datos formateados
  data.forEach(item => {
    // item.fechaEntrada y item.horaEntrada ya vienen strings limpios (dd/mm/yyyy HH:mm:ss)
    const fechaVisual = `${item.fechaEntrada} ${item.horaEntrada}`; 
    
    html += '<tr>';
    html += `<td><strong>${item.cedula}</strong></td>`;
    html += `<td>${item.nombre}</td>`;
    html += `<td>${item.centro}</td>`;
    html += `<td>${fechaVisual}</td>`; // Visualización limpia
    // Botón que llama al backend
    html += `<td><button class="btn-cerrar" onclick="google.script.run.withSuccessHandler(google.script.host.close).corregirDesdeSidebar('${item.cedula}')">🔒 Cerrar</button></td>`;
    html += '</tr>';
  });

  html += '</tbody></table></div>';

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setTitle('Corrección de Turnos')
    .setWidth(500);

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// ======================================================================
// 4. FUNCIÓN DE DATOS (Backend - Visual Limpia)
// ======================================================================

/**
 * @summary Obtiene la lista cruda de salidas pendientes.
 * @description
 *   - Lee la hoja 'Respuestas'.
 *   - Detecta filas donde la "Fecha Salida" está vacía.
 *   - **Lógica Robusta:** Si la fecha es "basura" (Ej: Jan 02 00:00), también la considera pendiente.
 *   - Aplica formateadores `spFormatearFecha_` y `spFormatearHora_` a la entrada (Q1).
 *   - Retorna array ordenado por fecha de entrada.
 * 
 * @returns {Array<Object>} Lista de pendientes listos para el Frontend.
 */
function getRawPendingList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SP_SHEET_RESPUESTAS);
  if (!hoja) return [];
  
  spAsegurarEncabezados_(hoja);
  const datos = hoja.getDataRange().getValues();
  if (!datos || datos.length <= 1) return [];

  const listaFinal = [];

  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    const ced = String(row[SP_I.CED] || '').trim();
    if (!ced) continue;

    // ---------------------------------------------------------
    // LÓGICA ROBUSTA DE "PENDIENTE"
    // ---------------------------------------------------------
    
    // 1. Si Fecha Salida está vacía (String "")
    let fechaSalStr = '';
    if (!row[SP_I.FECHA_SAL]) fechaSalStr = '';
    else if (row[SP_I.FECHA_SAL] instanceof Date) fechaSalStr = Utilities.formatDate(row[SP_I.FECHA_SAL], SP_TZ, "dd/MM/yyyy");
    else fechaSalStr = String(row[SP_I.FECHA_SAL]).trim();

    // 2. Detección de "Basura" (Fechas por defecto de celda vacía)
    // Si la celda quedó con valor por defecto (Fri Jan 02 00:00), se considera pendiente
    const esBasura = fechaSalStr.includes('Jan 02 00:00') || fechaSalStr.includes('Sun Dec 14 00:00');

    // 3. Decisión: Si está vacía O es basura, se añade a la lista
    if (!fechaSalStr || fechaSalStr === '' || esBasura) {
      
      // ✅ CORRECCIÓN Q1: Formatear fecha y hora de entrada correctamente
      const fechaFormateada = spFormatearFecha_(row[SP_I.FECHA_ENT]);
      const horaFormateada = spFormatearHora_(row[SP_I.HORA_ENT]);
      
      listaFinal.push({
        cedula: ced,
        nombre: String(row[SP_I.NOMBRE] || 'Desconocido'),
        centro: String(row[SP_I.CENTRO] || ''),
        fechaEntrada: fechaFormateada, // String dd/mm/yyyy
        horaEntrada: horaFormateada    // String HH:mm:ss
      });
    }
  }

  // Ordenar por fecha de entrada (más recientes arriba)
  // Convertimos dd/mm/yyyy a fecha comparable para sort
  listaFinal.sort((a, b) => {
    const parseF = (f) => {
      const p = f.split('/');
      return p.length === 3 ? new Date(`${p[2]}-${p[1]}-${p[0]}`) : new Date(0);
    };
    return parseF(b.fechaEntrada) - parseF(a.fechaEntrada);
  });

  return listaFinal;
}

// ======================================================================
// 5. FUNCIÓN DE CORRECCIÓN (Sidebar - Validación Manual)
// ======================================================================

/**
 * @summary Cierra un turno desde el Sidebar (Administración).
 * @description
 *   - Busca la última fila abierta de la cédula.
 *   - **Q2:** Pide fecha/hora en formato DD/MM/YYYY HH:MM (Manual).
 *   - Valida que la fecha sea POSTERIOR a la entrada.
 *   - Escribe los datos y realiza Auditoría.
 * 
 * @param {String} cedula - Cédula del empleado.
 */
function corregirDesdeSidebar(cedula) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SP_SHEET_RESPUESTAS);
  if (!hoja) { ui.alert("❌ Hoja no encontrada"); return; }

  spAsegurarEncabezados_(hoja);

  const ced = String(cedula).replace(/\D/g, '').trim();
  const finder = hoja.createTextFinder(ced).matchEntireCell(true);
  const hits = finder.findAll();
  
  if (!hits || hits.length === 0) {
    ui.alert("❌ Cédula no encontrada.");
    return;
  }

  // Buscar la ÚLTIMA fila sin salida (Turno más reciente abierto)
  let rowEntrada = -1;
  hits.sort((a, b) => b.getRow() - a.getRow());

  for (let i = 0; i < hits.length; i++) {
    const r = hits[i].getRow();
    if (r <= 1) continue;
    
    // Leemos Fecha Salida para verificar si está vacía
    const cSalida = String(hoja.getRange(r, SP_I.FECHA_SAL + 1).getValue() || '').trim();
    
    // Si NO tiene fecha de salida (está vacía), esta es la fila a cerrar
    if (!cSalida) { 
      rowEntrada = r; 
      break; 
    }
  }

  if (rowEntrada === -1) {
    ui.alert("❌ No se encontró entrada abierta (posiblemente se cerró mientras la lista estaba abierta).");
    return; 
  }

  // Obtener datos de entrada para mostrar al usuario
  const entradaVals = hoja.getRange(rowEntrada, 1, 1, hoja.getLastColumn()).getValues()[0];
  const nombre = String(entradaVals[SP_I.NOMBRE] || 'Sin nombre');
  const centro = String(entradaVals[SP_I.CENTRO] || '');
  
  // ✅ Formatear fecha/hora de entrada para mostrar en el prompt (Q1)
  const fechaEntradaStr = spFormatearFecha_(entradaVals[SP_I.FECHA_ENT]);
  const horaEntradaStr = spFormatearHora_(entradaVals[SP_I.HORA_ENT]);
  const entradaCompleta = `${fechaEntradaStr} ${horaEntradaStr}`;
  
  // Parsear para validación lógica (convertir dd/mm/yyyy HH:mm a objeto Date)
  const dtEntrada = spParsearFechaHora_(entradaCompleta) || new Date(0);
  
  // ✅ Q2: Generar sugerencia en formato DD/MM/YYYY HH:MM:SS
  const ahora = new Date();
  const sugerencia = Utilities.formatDate(ahora, SP_TZ, "dd/MM/yyyy HH:mm:ss");
  
  const mensaje = `Empleado: ${nombre}\nCédula: ${ced}\nCentro: ${centro}\nEntrada: ${entradaCompleta}\n\nIngresa fecha y hora de cierre:`;

  // ✅ Q2: Pedir en formato DD/MM/YYYY HH:MM:SS
  const fechaResp = ui.prompt(
    'FECHA Y HORA DE CIERRE',
    `${mensaje}\n\nFormato requerido: DD/MM/YYYY HH:MM:SS\nEjemplo: ${sugerencia}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (fechaResp.getSelectedButton() !== ui.Button.OK) return;

  const textoFecha = String(fechaResp.getResponseText() || '').trim();
  
  // ✅ Q2: Parsear formato DD/MM/YYYY HH:MM:SS manual
  const cierreDt = spParsearFechaHora_(textoFecha);

  if (!cierreDt) {
    ui.alert('❌ Formato inválido.\n\nUsa: DD/MM/YYYY HH:MM:SS\nEjemplo: 02/01/2026 14:30:00');
    return;
  }
  
  // ✅ Q2: Validación (Cierre > Entrada)
  if (dtEntrada.getTime() > 0 && cierreDt <= dtEntrada) {
    ui.alert(`❌ La fecha/hora de cierre debe ser POSTERIOR a la entrada.\n\nEntrada: ${entradaCompleta}\nCierre ingresado: ${textoFecha}`);
    return;
  }

  // Formatear para guardar en la hoja
  const fechaDDMMYYYY = Utilities.formatDate(cierreDt, SP_TZ, "dd/MM/yyyy");
  const horaHHMMSS = Utilities.formatDate(cierreDt, SP_TZ, "HH:mm:ss");

  // Escribir en la hoja
  spSetSalidaEnFilaEntrada_(hoja, rowEntrada, fechaDDMMYYYY, horaHHMMSS, '', '');
  
  // Auditoría
  if (SP_ADMIN_AUDIT_MODE === "sheet") {
    const audit = spEnsureAuditSheet_(ss);
    const usuario = (Session.getActiveUser && Session.getActiveUser().getEmail) ? (Session.getActiveUser().getEmail() || 'Desconocido') : 'ScriptUser';
    const obsActual = String(hoja.getRange(rowEntrada, SP_I.OBS + 1).getValue() || '');
    const obsFinal = obsActual ? `[ADMIN] ${obsActual}` : 'Cerrado desde Sidebar';
    hoja.getRange(rowEntrada, SP_I.OBS + 1).setValue(obsFinal);

    audit.appendRow([
      new Date(),
      ced,
      nombre,
      centro,
      rowEntrada,
      fechaDDMMYYYY,
      horaHHMMSS,
      "Cerrado desde Sidebar Admin",
      usuario
    ]);
  }

  ui.alert(`✅ Turno cerrado correctamente.\n\nCédula: ${ced}\nEmpleado: ${nombre}\nFila: ${rowEntrada}\nCierre: ${fechaDDMMYYYY} ${horaHHMMSS}`);
}

// ======================================================================
// 6. UTILIDAD DE ESCRITURA
// ======================================================================

/**
 * @summary Escribe los datos de cierre en la fila de entrada.
 * @private
 */
function spSetSalidaEnFilaEntrada_(hoja, rowEntrada, fechaDDMMYYYY, horaHHMMSS, fotoSalida, dentroSalida) {
  const colFechaSal = SP_I.FECHA_SAL + 1;
  const colHoraSal = SP_I.HORA_SAL + 1;
  const colFotoSal = SP_I.FOTO_SAL + 1;
  const colDentroSal = SP_I.DENTRO_SAL + 1;

  hoja.getRange(rowEntrada, colFechaSal).setValue(fechaDDMMYYYY);
  hoja.getRange(rowEntrada, colHoraSal).setValue(horaHHMMSS);
  hoja.getRange(rowEntrada, colFotoSal).setValue(fotoSalida || '');
  hoja.getRange(rowEntrada, colDentroSal).setValue(dentroSalida || '');
}

// ======================================================================
// 7. AUDITORÍA
// ======================================================================

/**
 * @summary Asegura que la hoja de Auditoría exista.
 * @private
 */
function spEnsureAuditSheet_(ss) {
  if (SP_ADMIN_AUDIT_MODE !== "sheet") return null; 
  let sh = ss.getSheetByName(SP_AUDIT_SHEET);
  if (!sh) sh = ss.insertSheet(SP_AUDIT_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "TimestampAdmin", "Cedula", "Nombre", "Centro",
      "FilaEntradaCerrada", "FechaSalida", "HoraSalida",
      "Motivo", "UsuarioAdmin"
    ]);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f3f3f3');
  }
  return sh;
}
