/**
 * ============================================================================
 *  SINCRONIZADOR DE PLANIFICACIÓN → GOOGLE CALENDAR (DII)
 * ============================================================================
 *  Lee la hoja activa de la planilla y crea/actualiza eventos en un Google
 *  Calendar compartido, evitando duplicados en re-ejecuciones.
 *
 *  Estructura esperada (fila 1 = encabezado, datos desde fila 2):
 *   A: Fecha            B: Acción / Actividad     C: Responsable
 *   D: Seguimiento      E: Relación acreditación  F: Tipo (Interna/Académica)
 *   G: Lugar de origen  H: Comentarios
 *
 *  Instalación:
 *   1. Extensiones > Apps Script en tu planilla de Google Sheets.
 *   2. Pega este archivo completo (reemplaza el contenido de Code.gs).
 *   3. Reemplaza CALENDAR_ID más abajo por el ID real del calendario.
 *   4. Guarda y recarga la planilla. Aparecerá el menú "Sincronizar
 *      Calendario". La primera ejecución pedirá autorización.
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// 1. CONFIGURACIÓN
// ----------------------------------------------------------------------------

/** ID del Google Calendar compartido donde se crearán los eventos.
 *  Se obtiene en Configuración del calendario > "ID de calendario". */
const CALENDAR_ID = 'REEMPLAZAR_CON_ID_DEL_CALENDARIO@group.calendar.google.com';

/** Mapa de columnas (1 = A, 2 = B, ...) para que el resto del código sea
 *  legible y fácil de mantener si cambia el orden de las columnas. */
const COL = {
  FECHA: 1,
  ACCION: 2,
  RESPONSABLE: 3,
  SEGUIMIENTO: 4,
  RELACION_ACREDITACION: 5,
  TIPO: 6,
  LUGAR_ORIGEN: 7,
  COMENTARIOS: 8,
};

const FILA_INICIO_DATOS = 2; // fila 1 = encabezado
const NUM_COLUMNAS = 8; // A..H

/** Color de evento según la columna "Tipo".
 *  Ver CalendarApp.EventColor para más opciones. */
const COLOR_POR_TIPO = {
  'Académica': CalendarApp.EventColor.PALE_BLUE,
  'Interna': CalendarApp.EventColor.YELLOW,
};

// ----------------------------------------------------------------------------
// 2. MENÚ EN GOOGLE SHEETS
// ----------------------------------------------------------------------------

/**
 * Se ejecuta automáticamente al abrir la planilla.
 * Agrega el menú personalizado para lanzar la sincronización manualmente.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sincronizar Calendario')
    .addItem('Ejecutar Sincronización', 'sincronizarCalendario')
    .addToUi();
}

// ----------------------------------------------------------------------------
// 3. FUNCIÓN PRINCIPAL
// ----------------------------------------------------------------------------

/**
 * Recorre la hoja activa, crea los eventos correspondientes en el calendario
 * y muestra un resumen al usuario al finalizar.
 */
function sincronizarCalendario() {
  const ui = SpreadsheetApp.getUi();
  let calendar;

  try {
    calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      throw new Error(
        'No se encontró el calendario. Verifica que CALENDAR_ID sea correcto ' +
        'y que la cuenta que ejecuta el script tenga acceso a él.'
      );
    }
  } catch (err) {
    ui.alert('Error de configuración', err.message, ui.ButtonSet.OK);
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  // Contadores para el resumen final.
  const resumen = {
    creados: 0,
    duplicados: 0,
    pendientes: 0,
    filasVacias: 0,
  };
  const detalleFilasPendientes = [];

  if (lastRow < FILA_INICIO_DATOS) {
    ui.alert('Sin datos', 'La hoja activa no tiene filas de datos para procesar.', ui.ButtonSet.OK);
    return;
  }

  const datos = sheet
    .getRange(FILA_INICIO_DATOS, 1, lastRow - FILA_INICIO_DATOS + 1, NUM_COLUMNAS)
    .getValues();

  datos.forEach((fila, indice) => {
    const numeroFila = indice + FILA_INICIO_DATOS; // fila real en la hoja

    const accion = normalizarTexto(fila[COL.ACCION - 1]);
    const fechaCruda = fila[COL.FECHA - 1];

    // Fila vacía o sin título: se omite silenciosamente.
    if (!accion && esVacio(fechaCruda)) {
      resumen.filasVacias++;
      return;
    }
    if (!accion) {
      resumen.filasVacias++;
      return;
    }

    // --- Parseo de fecha -----------------------------------------------
    const fechaInfo = parsearFecha(fechaCruda);

    if (!fechaInfo) {
      resumen.pendientes++;
      detalleFilasPendientes.push(`Fila ${numeroFila}: "${fechaCruda}"`);
      return;
    }

    // --- Armado de datos del evento --------------------------------------
    const tipo = normalizarTexto(fila[COL.TIPO - 1]);
    const color = COLOR_POR_TIPO[tipo] || null; // sin color = color por defecto del calendario
    const descripcion = construirDescripcion({
      responsable: normalizarTexto(fila[COL.RESPONSABLE - 1]),
      seguimiento: normalizarTexto(fila[COL.SEGUIMIENTO - 1]),
      relacion: normalizarTexto(fila[COL.RELACION_ACREDITACION - 1]),
      tipo: tipo,
      lugar: normalizarTexto(fila[COL.LUGAR_ORIGEN - 1]),
      comentarios: normalizarTexto(fila[COL.COMENTARIOS - 1]),
      fila: numeroFila,
    });

    // --- Prevención de duplicados + creación -----------------------------
    const yaExiste = eventoYaExiste(calendar, accion, fechaInfo);

    if (yaExiste) {
      resumen.duplicados++;
      return;
    }

    try {
      crearEvento(calendar, accion, fechaInfo, descripcion, color);
      resumen.creados++;
    } catch (err) {
      // Un error puntual en una fila no debe detener el resto de la sincronización.
      Logger.log(`Error creando evento en fila ${numeroFila}: ${err.message}`);
      resumen.pendientes++;
      detalleFilasPendientes.push(`Fila ${numeroFila}: error al crear evento (${err.message})`);
    }
  });

  mostrarResumen(ui, resumen, detalleFilasPendientes);
}

// ----------------------------------------------------------------------------
// 4. PARSEO DE FECHAS Y RANGOS
// ----------------------------------------------------------------------------

/**
 * Interpreta el contenido de la columna Fecha en sus distintas variantes
 * reales (fecha única como objeto Date, texto "dd/mm/aaaa", rangos con
 * distintos separadores y niveles de año/mes implícito, o texto ambiguo).
 *
 * @param {*} valorCrudo Valor tal cual viene de getValues().
 * @return {?{esRango: boolean, inicio: Date, finInclusive: Date}} null si la
 *   fecha es ambigua / no se pudo interpretar con certeza.
 */
function parsearFecha(valorCrudo) {
  if (esVacio(valorCrudo)) return null;

  // Caso 1: la celda ya viene formateada como fecha en Sheets -> objeto Date.
  if (Object.prototype.toString.call(valorCrudo) === '[object Date]') {
    return { esRango: false, inicio: valorCrudo, finInclusive: null };
  }

  const texto = normalizarTexto(valorCrudo);
  if (!texto) return null;

  // Los rangos en la planilla usan indistintamente "-" (guion) o "–" (guion
  // largo/en-dash), con o sin espacios alrededor. [-–] cubre ambos casos.

  // Caso 2: rango completo con separador de fecha en guiones, ej:
  //   "05-10-2026- 07-12-2026"  ->  05/10/2026 al 07/12/2026
  let m = texto.match(/^(\d{1,2})-(\d{1,2})-(\d{4})\s*-\s*(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) {
    const inicio = crearFechaLocal(m[3], m[2], m[1]);
    const fin = crearFechaLocal(m[6], m[5], m[4]);
    return { esRango: true, inicio, finInclusive: fin };
  }

  // Caso 3: rango con fecha completa dd/mm/aaaa a ambos lados, ej:
  //   "30/06/2026 – 17/07/2026", "27/07/2026– 31/07/2026"
  m = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s*[-–]\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const inicio = crearFechaLocal(m[3], m[2], m[1]);
    const fin = crearFechaLocal(m[6], m[5], m[4]);
    return { esRango: true, inicio, finInclusive: fin };
  }

  // Caso 4: inicio "dd/mm" sin año + fin "dd/mm/aaaa", ej: "30/11–11/12/2026".
  // El año del inicio se infiere del fin (los rangos de la planilla no cruzan
  // año en este formato; el caso que sí cruza año usa el formato del Caso 3).
  m = texto.match(/^(\d{1,2})\/(\d{1,2})\s*[-–]\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const anio = m[5];
    const inicio = crearFechaLocal(anio, m[2], m[1]);
    const fin = crearFechaLocal(anio, m[4], m[3]);
    return { esRango: true, inicio, finInclusive: fin };
  }

  // Caso 5: inicio solo "dd" + fin "dd/mm/aaaa", comparten mes/año, ej:
  //   "07–08/07/2026", "25–29/01/2027".
  m = texto.match(/^(\d{1,2})\s*[-–]\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const mes = m[3];
    const anio = m[4];
    const inicio = crearFechaLocal(anio, mes, m[1]);
    const fin = crearFechaLocal(anio, mes, m[2]);
    return { esRango: true, inicio, finInclusive: fin };
  }

  // Caso 6: fecha límite tipo "Hasta 28/11/2026" -> se registra como evento
  // de un solo día en la fecha indicada (el día del plazo/vencimiento).
  m = texto.match(/^Hasta\s+(\d{1,2})\/(\d{1,2})\/(\d{4})$/i);
  if (m) {
    const inicio = crearFechaLocal(m[3], m[2], m[1]);
    return { esRango: false, inicio, finInclusive: null };
  }

  // Caso 7: fecha única "dd/mm/aaaa" escrita como texto.
  m = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const inicio = crearFechaLocal(m[3], m[2], m[1]);
    return { esRango: false, inicio, finInclusive: null };
  }

  // Caso 8: cualquier otra cosa ("Por definir", "Agosto", "Octubre–Noviembre",
  // "Primera semana de agosto", etc.) se considera ambigua y se omite.
  return null;
}

/**
 * Crea un objeto Date a las 00:00 hora local a partir de componentes
 * separados, evitando ambigüedades de parseo de strings tipo "dd/mm/aaaa".
 */
function crearFechaLocal(anio, mes, dia) {
  return new Date(Number(anio), Number(mes) - 1, Number(dia));
}

// ----------------------------------------------------------------------------
// 5. PREVENCIÓN DE DUPLICADOS Y CREACIÓN DE EVENTOS
// ----------------------------------------------------------------------------

/**
 * Revisa si ya existe un evento con el mismo título en el día (o rango) dado,
 * para no volver a crearlo en re-ejecuciones del script.
 */
function eventoYaExiste(calendar, titulo, fechaInfo) {
  let eventosExistentes;

  if (!fechaInfo.esRango) {
    eventosExistentes = calendar.getEventsForDay(fechaInfo.inicio);
  } else {
    const finExclusivo = sumarDias(fechaInfo.finInclusive, 1);
    eventosExistentes = calendar.getEvents(fechaInfo.inicio, finExclusivo);
  }

  return eventosExistentes.some((ev) => ev.getTitle().trim() === titulo);
}

/**
 * Crea el evento de día único o de rango según corresponda, aplicando color
 * y descripción.
 */
function crearEvento(calendar, titulo, fechaInfo, descripcion, color) {
  let evento;

  if (!fechaInfo.esRango) {
    evento = calendar.createAllDayEvent(titulo, fechaInfo.inicio, {
      description: descripcion,
    });
  } else {
    // createAllDayEvent usa fecha de fin EXCLUSIVA, por eso se suma 1 día
    // a la fecha de fin inclusiva que trae la planilla.
    const finExclusivo = sumarDias(fechaInfo.finInclusive, 1);
    evento = calendar.createAllDayEvent(titulo, fechaInfo.inicio, finExclusivo, {
      description: descripcion,
    });
  }

  if (color) {
    evento.setColor(color);
  }

  return evento;
}

/** Devuelve una nueva fecha sumando N días a la fecha dada (no muta la original). */
function sumarDias(fecha, dias) {
  const nueva = new Date(fecha);
  nueva.setDate(nueva.getDate() + dias);
  return nueva;
}

// ----------------------------------------------------------------------------
// 6. DESCRIPCIÓN DEL EVENTO
// ----------------------------------------------------------------------------

/**
 * Arma el cuerpo de descripción del evento en Calendar, incluyendo solo los
 * campos que efectivamente tienen contenido.
 */
function construirDescripcion(campos) {
  const lineas = [];

  if (campos.responsable) lineas.push(`Responsable: ${campos.responsable}`);
  if (campos.seguimiento) lineas.push(`Seguimiento / Evidencia: ${campos.seguimiento}`);
  if (campos.relacion) lineas.push(`Relación con acreditación: ${campos.relacion}`);
  if (campos.tipo) lineas.push(`Tipo: ${campos.tipo}`);
  if (campos.lugar) lineas.push(`Lugar de origen: ${campos.lugar}`);
  if (campos.comentarios) lineas.push(`Comentarios: ${campos.comentarios}`);

  lineas.push('');
  lineas.push(`— Sincronizado automáticamente desde la planilla (fila ${campos.fila}).`);

  return lineas.join('\n');
}

// ----------------------------------------------------------------------------
// 7. UTILIDADES GENERALES
// ----------------------------------------------------------------------------

/** true si el valor es null, undefined o string vacío/solo espacios. */
function esVacio(valor) {
  if (valor === null || valor === undefined) return true;
  if (typeof valor === 'string' && valor.trim() === '') return true;
  return false;
}

/** Convierte a string, recorta espacios y colapsa espacios internos múltiples. */
function normalizarTexto(valor) {
  if (esVacio(valor)) return '';
  return String(valor).trim().replace(/\s+/g, ' ');
}

// ----------------------------------------------------------------------------
// 8. RESUMEN FINAL
// ----------------------------------------------------------------------------

/** Muestra el alert() con el resumen de la sincronización. */
function mostrarResumen(ui, resumen, detalleFilasPendientes) {
  let mensaje =
    `Eventos creados: ${resumen.creados}\n` +
    `Omitidos por duplicado: ${resumen.duplicados}\n` +
    `Ignorados por fecha pendiente/inválida: ${resumen.pendientes}\n`;

  if (resumen.filasVacias > 0) {
    mensaje += `\n(${resumen.filasVacias} fila(s) vacías o sin título fueron omitidas.)`;
  }

  if (detalleFilasPendientes.length > 0) {
    // Se muestra un máximo de 10 para no saturar el alert.
    const detalle = detalleFilasPendientes.slice(0, 10).join('\n');
    const extra = detalleFilasPendientes.length > 10
      ? `\n... y ${detalleFilasPendientes.length - 10} más (ver Ejecuciones/Logs).`
      : '';
    mensaje += `\n\nDetalle de filas pendientes:\n${detalle}${extra}`;
    Logger.log('Filas pendientes:\n' + detalleFilasPendientes.join('\n'));
  }

  ui.alert('Resumen de Sincronización', mensaje, ui.ButtonSet.OK);
}