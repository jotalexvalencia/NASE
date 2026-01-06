/**
 * ============================================================
 * ⚙️ ConfigHorarios.gs – Gestor de Configuración de Tiempos (NASE 2026)
 * ======================================================================
 * @summary Módulo de almacenamiento de preferencias horarias.
 * @description Este archivo actúa como el "almacén de configuración" del sistema.
 *              Anteriormente (Ley 2025) se usaba para guardar porcentajes de recargos
 *              monetarios. Se ha simplificado (LIMPIEZA) y ahora se dedica exclusivamente
 *              a guardar los rangos de horarios nocturnos.
 *
 * @features
 *   - 🧹 **Limpieza de Código:** Se eliminaron todas las propiedades y funciones relacionadas
 *     con recargos monetarios (%) porque el sistema actual los calcula de otra forma.
 *   - 💾 **Persistencia:** Usa `ScriptProperties` (almacén clave-valor de Apps Script)
 *     para guardar las preferencias. Es más rápido y seguro que escribir en una hoja de cálculo.
 *   - 🕰️ **Configuración por Defecto:** Ley 2025 define Nocturno como 19:00 a 06:00 (7 PM a 6 AM).
 *     Este archivo maneja la lectura y escritura de esos valores.
 *
 * @dependencies
 *   - `hoja_turnos.gs`: Llama a `obtenerConfiguracionHorarios()` para saber cuándo
 *     cuenta una hora como nocturna.
 *   - `config_horarios.html` (si existe): Llama a `actualizarConfiguracionHorarios()` para guardar cambios.
 *
 * @author NASE Team
 * @version 2.0 (Simplificado - Solo Tiempos)
 */

// ======================================================================
// 1. CONSTANTES (Nombres de Propiedades)
// ======================================================================

const CONFIG_PROPS = {
  // Claves en `ScriptProperties` para guardar las horas de inicio/fin nocturno
  HORA_NOCTURNA_INICIO: 'HORA_NOCTURNA_INICIO', // Default: 19 (7 PM)
  HORA_NOCTURNA_FIN: 'HORA_NOCTURNA_FIN',         // Default: 6  (6 AM)
  
  // Nota: Se eliminaron constantes como 'RECARGO_NOCTURNO' o 'RECARGO_FESTIVO'
  // porque el cálculo monetario ya no se gestiona desde aquí.
};

// ======================================================================
// 2. LECTURA DE CONFIGURACIÓN
// ======================================================================

/**
 * @summary Obtiene la configuración actual de horarios.
 * @description Lee las propiedades de `ScriptProperties`.
 *              Si no existen (primera ejecución), devuelve los valores por defecto
 *              establecidos por la Ley 2025 (19:00 - 06:00).
 * 
 * @returns {Object} Objeto con:
 *   - `horaInicio` (Number): Hora de inicio del recargo nocturno (Ej: 19).
 *   - `horaFin` (Number): Hora de fin del recargo nocturno (Ej: 6).
 */
function obtenerConfiguracionHorarios() {
  const props = PropertiesService.getScriptProperties();
  
  // Se leen solo las horas. Si la propiedad no existe en la memoria del script,
  // se usa el valor por defecto (19 y 6).
  return {
    horaInicio: parseInt(props.getProperty(CONFIG_PROPS.HORA_NOCTURNA_INICIO) || '19', 10),
    horaFin: parseInt(props.getProperty(CONFIG_PROPS.HORA_NOCTURNA_FIN) || '6', 10)
  };
}

// ======================================================================
// 3. ESCRITURA DE CONFIGURACIÓN
// ======================================================================

/**
 * @summary Actualiza (Guarda) la configuración de horarios.
 * @description Se ejecuta desde el formulario de configuración HTML.
 *              Guarda las horas de inicio y fin en `ScriptProperties` para que
 *              persistan entre ejecuciones.
 * 
 * @param {Object} config - Objeto con:
 *   - `horaInicio` (Number): Nueva hora de inicio (0-23).
 *   - `horaFin` (Number): Nueva hora de fin (0-23).
 * 
 * @returns {Object} { status: 'ok', message: String }
 */
function actualizarConfiguracionHorarios(config) {
  const props = PropertiesService.getScriptProperties();
  
  // Guardar hora de inicio si se proporcionó
  if (config.horaInicio !== undefined) {
    props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_INICIO, String(config.horaInicio));
  }
  
  // Guardar hora de fin si se proporcionó
  if (config.horaFin !== undefined) {
    props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_FIN, String(config.horaFin));
  }
  
  // Nota: Ya no guardamos porcentajes monetarios aquí (Limpieza de código).
  
  return { status: 'ok', message: 'Configuración de horarios actualizada correctamente.' };
}

// ======================================================================
// 4. RESET DE CONFIGURACIÓN
// ======================================================================

/**
 * @summary Restablece los valores por defecto (Ley 2025).
 * @description Función de seguridad para volver al estado original del sistema.
 *              Borra las propiedades personalizadas y fuerza el uso de 19:00 - 06:00.
 * 
 * @returns {Object} { status: 'ok', message: String }
 */
function restablecerConfiguracionPorDefecto() {
  const props = PropertiesService.getScriptProperties();
  
  // Sobrescribir con valores por defecto (19 y 6)
  props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_INICIO, '19');
  props.setProperty(CONFIG_PROPS.HORA_NOCTURNA_FIN, '6');
  
  return { status: 'ok', message: 'Horarios restablecidos a ley 2025 (19:00 - 06:00).' };
}
