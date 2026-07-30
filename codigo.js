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
 *   G: Lugar de origen  H: Comentarios             I: ID Evento (uso interno,
 *                                                      no editar manualmente)
 *
 *  La columna I la escribe el propio script para recordar qué evento de
 *  Calendar corresponde a cada fila. Con eso se logra que:
 *   - Si se modifica cualquier dato de una fila (incluida la fecha), el
 *     evento existente se actualiza en vez de crear uno nuevo.
 *   - Si una fila se elimina de la planilla, su evento en Calendar también
 *     se elimina automáticamente en la siguiente sincronización.
 *
 *  Diseño orientado a rendimiento: en vez de consultar el calendario una vez
 *  por cada fila de la planilla, el script trae en una sola consulta todos
 *  los eventos del periodo cubierto por los datos, arma con eso un índice en
 *  memoria, y solo crea o modifica un evento cuando el contenido de la fila
 *  realmente difiere del evento existente. Esto reduce el número de llamadas
 *  a servicios externos (la parte lenta de cualquier script de Apps Script)
 *  a un puñado por ejecución, en vez de varias por fila.
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
  ID_EVENTO: 9, // Columna donde el script guarda el ID del evento de Calendar
};

const FILA_INICIO_DATOS = 2; // fila 1 = encabezado
const NUM_COLUMNAS = 9; // A..I

/** Clave usada en las propiedades del documento para recordar, entre
 *  ejecuciones, qué IDs de evento fueron sincronizados por el script. Esto
 *  es lo que permite detectar filas eliminadas y borrar su evento asociado. */
const PROP_IDS_SINCRONIZADOS = 'EVENTOS_SINCRONIZADOS';

/** Color de evento según la columna "Tipo".
 *  Ver CalendarApp.EventColor para más opciones. */
const COLOR_POR_TIPO = {
  'Académica': CalendarApp.EventColor.PALE_BLUE,
  'Interna': CalendarApp.EventColor.YELLOW,
};

/** Margen (en días) que se suma alrededor del rango de fechas de la
 *  planilla al consultar el calendario, para no dejar fuera eventos cuya
 *  fecha se haya corrido levemente respecto a la última sincronización. */
const MARGEN_DIAS_CONSULTA = 7;

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
 * Recorre la hoja activa, crea/actualiza los eventos correspondientes en el
 * calendario, elimina los eventos de filas que ya no existen y muestra un
 * resumen al usuario al finalizar.
 */
function sincronizarCalendario() {
  const marcaInicio = Date.now();
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

  // Asegura el encabezado de la columna interna de control.
  if (esVacio(sheet.getRange(1, COL.ID_EVENTO).getValue())) {
    sheet.getRange(1, COL.ID_EVENTO).setValue('ID Evento');
  }

  if (lastRow < FILA_INICIO_DATOS) {
    ui.alert('Sin datos', 'La hoja activa no tiene filas de datos para procesar.', ui.ButtonSet.OK);
    return;
  }

  // Contadores para el resumen final.
  const resumen = {
    creados: 0,
    actualizados: 0,
    sinCambios: 0,
    eliminados: 0,
    pendientes: 0,
    filasVacias: 0,
  };
  const detalleFilasPendientes = [];

  // Registro persistente de IDs sincronizados en ejecuciones previas, usado
  // para detectar filas eliminadas (reconciliación).
  const propiedades = PropertiesService.getDocumentProperties();
  const idsPrevios = JSON.parse(propiedades.getProperty(PROP_IDS_SINCRONIZADOS) || '[]');
  const idsVistos = new Set(); // IDs que siguen existiendo en esta ejecución

  const datos = sheet
    .getRange(FILA_INICIO_DATOS, 1, lastRow - FILA_INICIO_DATOS + 1, NUM_COLUMNAS)
    .getValues();

  // Valor por defecto de la columna ID Evento para cada fila: el que ya
  // tenía. Las filas que se procesen más abajo lo sobrescriben; las que se
  // omiten (vacías o con fecha pendiente) simplemente lo conservan.
  const columnaIdEvento = datos.map((fila) => [normalizarTexto(fila[COL.ID_EVENTO - 1])]);

  // --- Primera pasada: se interpreta cada fila y se determina el rango de
  // fechas cubierto por la planilla, sin todavía consultar el calendario. ---
  const filasProcesables = [];
  let minFecha = null;
  let maxFecha = null;

  datos.forEach((fila, indice) => {
    const numeroFila = indice + FILA_INICIO_DATOS;
    const accion = normalizarTexto(fila[COL.ACCION - 1]);
    const fechaCruda = fila[COL.FECHA - 1];

    // Fila sin título: se omite. Si tenía un evento asociado, su ID no se
    // agregará a idsVistos y por lo tanto será eliminado en la etapa de
    // eventos huérfanos, igual que si la fila se hubiera borrado.
    if (!accion) {
      resumen.filasVacias++;
      return;
    }

    const fechaInfo = parsearFecha(fechaCruda);
    if (!fechaInfo) {
      resumen.pendientes++;
      detalleFilasPendientes.push(`Fila ${numeroFila}: "${fechaCruda}"`);
      return;
    }

    const finRelevante = fechaInfo.esRango ? fechaInfo.finInclusive : fechaInfo.inicio;
    if (minFecha === null || fechaInfo.inicio < minFecha) minFecha = fechaInfo.inicio;
    if (maxFecha === null || finRelevante > maxFecha) maxFecha = finRelevante;

    filasProcesables.push({ indice, numeroFila, fila, accion, fechaInfo });
  });

  // --- Se traen en una sola consulta todos los eventos del calendario que
  // caen dentro del periodo cubierto por la planilla. A partir de esa única
  // respuesta se arman dos índices en memoria (por ID de evento y por día)
  // que reemplazan las consultas que antes se hacían al calendario una por
  // una, fila por fila. ---------------------------------------------------
  const eventosPorId = new Map();
  const eventosPorDia = new Map();

  if (filasProcesables.length > 0) {
    const timeMin = sumarDias(minFecha, -MARGEN_DIAS_CONSULTA);
    const timeMax = sumarDias(maxFecha, MARGEN_DIAS_CONSULTA + 1);
    calendar.getEvents(timeMin, timeMax).forEach((evento) => {
      registrarEventoEnIndices(evento, eventosPorId, eventosPorDia);
    });
  }

  // --- Segunda pasada: para cada fila se decide si corresponde crear el
  // evento, actualizarlo (solo si algún dato cambió respecto al existente) o
  // dejarlo tal como está. ---------------------------------------------------
  filasProcesables.forEach(({ indice, numeroFila, fila, accion, fechaInfo }) => {
    const tipo = normalizarTexto(fila[COL.TIPO - 1]);
    const color = COLOR_POR_TIPO[tipo] || null;
    const descripcion = construirDescripcion({
      responsable: normalizarTexto(fila[COL.RESPONSABLE - 1]),
      seguimiento: normalizarTexto(fila[COL.SEGUIMIENTO - 1]),
      relacion: normalizarTexto(fila[COL.RELACION_ACREDITACION - 1]),
      tipo: tipo,
      lugar: normalizarTexto(fila[COL.LUGAR_ORIGEN - 1]),
      comentarios: normalizarTexto(fila[COL.COMENTARIOS - 1]),
      fila: numeroFila,
    });

    const idGuardado = normalizarTexto(fila[COL.ID_EVENTO - 1]);

    try {
      const existente = obtenerEventoExistente(calendar, eventosPorId, eventosPorDia, idGuardado, accion, fechaInfo);
      let evento;

      if (existente) {
        const huboCambios = actualizarEventoSiCorresponde(existente, accion, fechaInfo, descripcion, color);
        evento = existente;
        if (huboCambios) {
          resumen.actualizados++;
        } else {
          resumen.sinCambios++;
        }
      } else {
        evento = crearEvento(calendar, accion, fechaInfo, descripcion, color);
        resumen.creados++;
      }

      registrarEventoEnIndices(evento, eventosPorId, eventosPorDia);
      idsVistos.add(evento.getId());
      columnaIdEvento[indice][0] = evento.getId();
    } catch (err) {
      // Un error puntual en una fila no debe detener el resto de la sincronización.
      Logger.log(`Error creando/actualizando evento en fila ${numeroFila}: ${err.message}`);
      resumen.pendientes++;
      detalleFilasPendientes.push(`Fila ${numeroFila}: error al sincronizar evento (${err.message})`);
    }
  });

  // --- Eliminación de eventos huérfanos (filas borradas o vaciadas) -------
  // Se reutiliza el índice ya cargado en memoria; solo se consulta el
  // calendario individualmente si el evento no formaba parte del periodo
  // recién traído (por ejemplo, si su fecha original quedó fuera del rango
  // cubierto por los datos actuales de la planilla).
  idsPrevios.forEach((id) => {
    if (idsVistos.has(id)) return;
    try {
      const eventoHuerfano = eventosPorId.has(id) ? eventosPorId.get(id) : calendar.getEventById(id);
      if (eventoHuerfano) {
        eventoHuerfano.deleteEvent();
        resumen.eliminados++;
      }
    } catch (err) {
      Logger.log(`No se pudo eliminar el evento huérfano ${id}: ${err.message}`);
    }
  });

  // Persiste el estado actual para poder detectar eliminaciones la próxima vez.
  propiedades.setProperty(PROP_IDS_SINCRONIZADOS, JSON.stringify(Array.from(idsVistos)));

  // Todos los ID de evento se escriben en una sola operación, en vez de una
  // escritura por fila, para minimizar las llamadas a la hoja de cálculo.
  sheet
    .getRange(FILA_INICIO_DATOS, COL.ID_EVENTO, columnaIdEvento.length, 1)
    .setValues(columnaIdEvento);

  const segundosTranscurridos = ((Date.now() - marcaInicio) / 1000).toFixed(1);
  mostrarResumen(ui, resumen, detalleFilasPendientes, segundosTranscurridos);
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
// 5. ÍNDICES DE EVENTOS Y OPERACIONES SOBRE EL CALENDARIO
// ----------------------------------------------------------------------------

/**
 * Incorpora un evento a los dos índices en memoria: por ID (búsqueda directa
 * cuando la fila ya tiene un ID Evento guardado) y por día (respaldo para
 * filas que todavía no tienen ID guardado, buscando por título y fecha).
 */
function registrarEventoEnIndices(evento, eventosPorId, eventosPorDia) {
  eventosPorId.set(evento.getId(), evento);

  const rango = obtenerRangoDeDias(evento);
  if (!rango) return;

  for (let dia = new Date(rango.inicio); dia < rango.fin; dia.setDate(dia.getDate() + 1)) {
    const clave = formatearClaveDia(dia);
    if (!eventosPorDia.has(clave)) eventosPorDia.set(clave, []);
    eventosPorDia.get(clave).push(evento);
  }
}

/**
 * Determina el rango de días (fin exclusivo) que ocupa un evento existente,
 * ya sea de día completo o con horario, para poder indexarlo día por día.
 */
function obtenerRangoDeDias(evento) {
  if (evento.isAllDayEvent()) {
    return { inicio: evento.getAllDayStartDate(), fin: evento.getAllDayEndDate() };
  }

  const inicio = new Date(evento.getStartTime());
  inicio.setHours(0, 0, 0, 0);
  const fin = new Date(evento.getEndTime());
  fin.setHours(0, 0, 0, 0);
  fin.setDate(fin.getDate() + 1);

  return { inicio, fin };
}

/**
 * Recupera el evento ya asociado a la fila. Primero por el ID guardado en la
 * columna "ID Evento" (mecanismo principal, inmune a cambios de título o
 * fecha), consultando primero el índice en memoria y, solo si no aparece ahí,
 * al calendario directamente. Si la fila todavía no tiene ID guardado (por
 * ejemplo, filas creadas con una versión anterior del script), se recurre a
 * la búsqueda por título y fecha para no duplicar eventos ya existentes.
 *
 * @return {?CalendarEvent} el evento existente, o null si no hay ninguno.
 */
function obtenerEventoExistente(calendar, eventosPorId, eventosPorDia, idGuardado, titulo, fechaInfo) {
  if (idGuardado) {
    if (eventosPorId.has(idGuardado)) return eventosPorId.get(idGuardado);
    const evento = calendar.getEventById(idGuardado);
    if (evento) return evento;
  }

  const candidatos = eventosPorDia.get(formatearClaveDia(fechaInfo.inicio)) || [];
  return candidatos.find((ev) => ev.getTitle().trim() === titulo) || null;
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

/**
 * Actualiza un evento existente para reflejar los datos actuales de la fila
 * (título, fecha/rango, descripción y color), pero solo llama a los métodos
 * de escritura de los campos que efectivamente cambiaron respecto al estado
 * actual del evento. Esto evita reescribir el evento completo en cada
 * sincronización cuando la fila no tuvo modificaciones.
 *
 * @return {boolean} true si se modificó algún campo del evento.
 */
function actualizarEventoSiCorresponde(evento, titulo, fechaInfo, descripcion, color) {
  let huboCambios = false;

  if (evento.getTitle() !== titulo) {
    evento.setTitle(titulo);
    huboCambios = true;
  }

  if (evento.getDescription() !== descripcion) {
    evento.setDescription(descripcion);
    huboCambios = true;
  }

  if (fechasDistintas(evento, fechaInfo)) {
    if (!fechaInfo.esRango) {
      evento.setAllDayDate(fechaInfo.inicio);
    } else {
      const finExclusivo = sumarDias(fechaInfo.finInclusive, 1);
      evento.setAllDayDates(fechaInfo.inicio, finExclusivo);
    }
    huboCambios = true;
  }

  if (color && String(evento.getColor()) !== String(color)) {
    evento.setColor(color);
    huboCambios = true;
  }

  return huboCambios;
}

/** true si la fecha/rango del evento no coincide con la que le corresponde según la planilla. */
function fechasDistintas(evento, fechaInfo) {
  if (!evento.isAllDayEvent()) return true;

  const finDeseadoExclusivo = sumarDias(fechaInfo.esRango ? fechaInfo.finInclusive : fechaInfo.inicio, 1);
  return (
    !mismaFecha(evento.getAllDayStartDate(), fechaInfo.inicio) ||
    !mismaFecha(evento.getAllDayEndDate(), finDeseadoExclusivo)
  );
}

/** Compara dos fechas por año, mes y día, ignorando la hora. */
function mismaFecha(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

/** Devuelve una nueva fecha sumando N días a la fecha dada (no muta la original). */
function sumarDias(fecha, dias) {
  const nueva = new Date(fecha);
  nueva.setDate(nueva.getDate() + dias);
  return nueva;
}

/** Genera una clave "aaaa-mm-dd" a partir de una fecha, usada como índice de día. */
function formatearClaveDia(fecha) {
  return fecha.getFullYear() + '-' + dosDigitos(fecha.getMonth() + 1) + '-' + dosDigitos(fecha.getDate());
}

function dosDigitos(numero) {
  return numero < 10 ? '0' + numero : String(numero);
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
function mostrarResumen(ui, resumen, detalleFilasPendientes, segundosTranscurridos) {
  let mensaje =
    `Eventos creados: ${resumen.creados}\n` +
    `Eventos actualizados: ${resumen.actualizados}\n` +
    `Eventos sin cambios: ${resumen.sinCambios}\n` +
    `Eventos eliminados (filas borradas): ${resumen.eliminados}\n` +
    `Filas con fecha pendiente o inválida: ${resumen.pendientes}\n` +
    `Tiempo de ejecución: ${segundosTranscurridos} segundos`;

  if (resumen.filasVacias > 0) {
    mensaje += `\n\n${resumen.filasVacias} fila(s) vacías o sin título fueron omitidas.`;
  }

  if (detalleFilasPendientes.length > 0) {
    // Se muestra un máximo de 10 para no saturar el alert.
    const detalle = detalleFilasPendientes.slice(0, 10).join('\n');
    const extra = detalleFilasPendientes.length > 10
      ? `\n... y ${detalleFilasPendientes.length - 10} más (ver Ejecuciones/Registros).`
      : '';
    mensaje += `\n\nDetalle de filas pendientes:\n${detalle}${extra}`;
    Logger.log('Filas pendientes:\n' + detalleFilasPendientes.join('\n'));
  }

  ui.alert('Resumen de Sincronización', mensaje, ui.ButtonSet.OK);
}