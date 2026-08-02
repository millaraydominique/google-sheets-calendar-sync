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
 *  Acceso multiusuario: la interacción con Calendar se hace a través del
 *  servicio avanzado "Google Calendar API" (identificador Calendar.*) y no
 *  a través del servicio básico CalendarApp. CalendarApp.getCalendarById()
 *  tiene una limitación conocida de Apps Script por la cual, para un
 *  calendario del que el usuario ejecutor no es dueño, exige de hecho un
 *  nivel de permiso más alto que "Hacer cambios en eventos" para reconocer
 *  el calendario, aunque ese nivel sí alcanza para leer, crear y editar
 *  eventos vía la API. El servicio avanzado respeta el nivel de permiso real
 *  otorgado al compartir el calendario, por lo que funciona igual sin
 *  importar qué cuenta ejecute el script, siempre que tenga al menos permiso
 *  de "Hacer cambios en eventos".
 *
 *  Instalación:
 *   1. Extensiones > Apps Script en tu planilla de Google Sheets.
 *   2. Pega este archivo completo (reemplaza el contenido de Code.gs).
 *   3. Reemplaza CALENDAR_ID más abajo por el ID real del calendario.
 *   4. En el editor de Apps Script, ve a "Servicios" (ícono +) y agrega
 *      "Google Calendar API". Debe quedar disponible con el identificador
 *      "Calendar". Este paso es obligatorio: sin él, el script no
 *      funcionará para ningún usuario, incluido el dueño del calendario.
 *   5. Guarda y recarga la planilla. Aparecerá el menú "Sincronizar
 *      Calendario". La primera ejecución de cada usuario pedirá
 *      autorización.
 *   6. Cada persona que vaya a ejecutar la sincronización debe tener el
 *      calendario compartido con su cuenta con permiso de, al menos,
 *      "Hacer cambios en eventos".
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
 *  Se usan las constantes de CalendarApp.EventColor porque son valores de
 *  solo texto (no requieren acceso al calendario) y coinciden exactamente
 *  con los valores numéricos de colorId que usa la Google Calendar API, por
 *  lo que se pueden reutilizar directamente. Ver CalendarApp.EventColor para
 *  más opciones. */
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

  try {
    Calendar.Calendars.get(CALENDAR_ID);
  } catch (err) {
    ui.alert(
      'Error de configuración',
      'No se encontró el calendario, o la cuenta que ejecuta el script no tiene acceso a él. ' +
        'Verifica que CALENDAR_ID sea correcto y que el calendario esté compartido con esta cuenta ' +
        'con permiso de al menos "Hacer cambios en eventos".\n\nDetalle técnico: ' + err.message,
      ui.ButtonSet.OK
    );
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
    listarEventosEnRango(CALENDAR_ID, timeMin, timeMax).forEach((evento) => {
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
      const existente = obtenerEventoExistente(CALENDAR_ID, eventosPorId, eventosPorDia, idGuardado, accion, fechaInfo);
      let evento;

      if (existente) {
        const resultado = actualizarEventoSiCorresponde(CALENDAR_ID, existente, accion, fechaInfo, descripcion, color);
        evento = resultado.evento;
        if (resultado.huboCambios) {
          resumen.actualizados++;
        } else {
          resumen.sinCambios++;
        }
      } else {
        evento = crearEvento(CALENDAR_ID, accion, fechaInfo, descripcion, color);
        resumen.creados++;
      }

      registrarEventoEnIndices(evento, eventosPorId, eventosPorDia);
      idsVistos.add(evento.id);
      columnaIdEvento[indice][0] = evento.id;
    } catch (err) {
      // Un error puntual en una fila no debe detener el resto de la sincronización.
      Logger.log(`Error creando/actualizando evento en fila ${numeroFila}: ${err.message}`);
      resumen.pendientes++;
      detalleFilasPendientes.push(`Fila ${numeroFila}: error al sincronizar evento (${err.message})`);
    }
  });

  // --- Eliminación de eventos huérfanos (filas borradas o vaciadas) -------
  idsPrevios.forEach((id) => {
    if (idsVistos.has(id)) return;
    try {
      Calendar.Events.remove(CALENDAR_ID, id);
      resumen.eliminados++;
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

/** Nombres de mes en español (en minúsculas, sin acentos) -> número de mes
 *  (1-12). Se acepta "setiembre" además de "septiembre" por ser una variante
 *  de uso común. */
const MESES_ES = {
  enero: 1,
  febrero: 2,
  marzo: 3,
  abril: 4,
  mayo: 5,
  junio: 6,
  julio: 7,
  agosto: 8,
  septiembre: 9,
  setiembre: 9,
  octubre: 10,
  noviembre: 11,
  diciembre: 12,
};

/** Rango de días [inicio, fin] para cada semana "ordinal" del mes usada en
 *  expresiones como "Primera semana de agosto". */
const SEMANAS_DEL_MES = {
  primera: [1, 7],
  segunda: [8, 14],
  tercera: [15, 21],
  cuarta: [22, 28],
};

/**
 * Interpreta el contenido de la columna Fecha en sus distintas variantes
 * reales (fecha única como objeto Date, texto "dd/mm/aaaa", rangos con
 * distintos separadores y niveles de año/mes implícito, expresiones en
 * lenguaje natural como "Octubre" o "Primera semana de agosto", o texto
 * ambiguo).
 *
 * @param {*} valorCrudo Valor tal cual viene de getValues().
 * @return {?{esRango: boolean, inicio: Date, finInclusive: Date}} null si la
 *   fecha es ambigua / no se pudo interpretar con certeza (por ejemplo "Por
 *   definir", "TBD", "Pendiente" o una celda vacía).
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

  // --- A partir de aquí: expresiones en lenguaje natural. Cuando la
  // expresión no menciona año, se asume el año actual. ------------------------

  // Caso 8: "Primera/Segunda/Tercera/Cuarta semana de [Mes]" [año opcional].
  m = texto.match(/^(Primera|Segunda|Tercera|Cuarta)\s+semana\s+de\s+([A-Za-zÀ-ÿ]+)(?:\s+(\d{4}))?$/i);
  if (m) {
    const mes = obtenerNumeroMes(m[2]);
    if (mes) {
      const anio = m[3] ? Number(m[3]) : new Date().getFullYear();
      const [diaInicio, diaFin] = SEMANAS_DEL_MES[m[1].toLowerCase()];
      const inicio = crearFechaLocal(anio, mes, diaInicio);
      const fin = crearFechaLocal(anio, mes, diaFin);
      return { esRango: true, inicio, finInclusive: fin };
    }
  }

  // Caso 9: "Mediados de [Mes]" [año opcional] -> días 11 al 20.
  m = texto.match(/^Mediados\s+de\s+([A-Za-zÀ-ÿ]+)(?:\s+(\d{4}))?$/i);
  if (m) {
    const mes = obtenerNumeroMes(m[1]);
    if (mes) {
      const anio = m[2] ? Number(m[2]) : new Date().getFullYear();
      const inicio = crearFechaLocal(anio, mes, 11);
      const fin = crearFechaLocal(anio, mes, 20);
      return { esRango: true, inicio, finInclusive: fin };
    }
  }

  // Caso 10: "Finales de [Mes]" / "Fin de [Mes]" [año opcional] -> día 21 al
  // último día del mes.
  m = texto.match(/^(Finales|Fin)\s+de\s+([A-Za-zÀ-ÿ]+)(?:\s+(\d{4}))?$/i);
  if (m) {
    const mes = obtenerNumeroMes(m[2]);
    if (mes) {
      const anio = m[3] ? Number(m[3]) : new Date().getFullYear();
      const inicio = crearFechaLocal(anio, mes, 21);
      const fin = crearFechaLocal(anio, mes, ultimoDiaDelMes(anio, mes));
      return { esRango: true, inicio, finInclusive: fin };
    }
  }

  // Caso 11: rango de meses completos, ej. "Octubre–Noviembre",
  // "Octubre-Noviembre 2026" -> desde el día 1 del primer mes hasta el
  // último día del segundo mes.
  m = texto.match(/^([A-Za-zÀ-ÿ]+)\s*[-–]\s*([A-Za-zÀ-ÿ]+)(?:\s+(\d{4}))?$/i);
  if (m) {
    const mesInicio = obtenerNumeroMes(m[1]);
    const mesFin = obtenerNumeroMes(m[2]);
    if (mesInicio && mesFin) {
      const anio = m[3] ? Number(m[3]) : new Date().getFullYear();
      const inicio = crearFechaLocal(anio, mesInicio, 1);
      const fin = crearFechaLocal(anio, mesFin, ultimoDiaDelMes(anio, mesFin));
      return { esRango: true, inicio, finInclusive: fin };
    }
  }

  // Caso 12: mes completo, ej. "Octubre", "Septiembre 2027" -> desde el día 1
  // hasta el último día de ese mes.
  m = texto.match(/^([A-Za-zÀ-ÿ]+)(?:\s+(\d{4}))?$/i);
  if (m) {
    const mes = obtenerNumeroMes(m[1]);
    if (mes) {
      const anio = m[2] ? Number(m[2]) : new Date().getFullYear();
      const inicio = crearFechaLocal(anio, mes, 1);
      const fin = crearFechaLocal(anio, mes, ultimoDiaDelMes(anio, mes));
      return { esRango: true, inicio, finInclusive: fin };
    }
  }

  // Caso 13: cualquier otra cosa ("Por definir", "TBD", "Pendiente", texto no
  // reconocido, etc.) se considera ambigua y se omite.
  return null;
}

/**
 * Crea un objeto Date a las 00:00 hora local a partir de componentes
 * separados, evitando ambigüedades de parseo de strings tipo "dd/mm/aaaa".
 */
function crearFechaLocal(anio, mes, dia) {
  return new Date(Number(anio), Number(mes) - 1, Number(dia));
}

/** Convierte un nombre de mes en español (cualquier combinación de
 *  mayúsculas/minúsculas) a su número de mes (1-12), o null si el texto no
 *  corresponde a un mes válido. */
function obtenerNumeroMes(nombre) {
  if (!nombre) return null;
  return MESES_ES[nombre.trim().toLowerCase()] || null;
}

/** Día del mes correspondiente al último día de {anio}-{mes} (mes 1-indexado,
 *  igual que en el resto del script: enero = 1, diciembre = 12). */
function ultimoDiaDelMes(anio, mes) {
  return new Date(Number(anio), Number(mes), 0).getDate();
}

// ----------------------------------------------------------------------------
// 5. ÍNDICES DE EVENTOS Y OPERACIONES SOBRE EL CALENDARIO
// ----------------------------------------------------------------------------
//  Esta sección usa el servicio avanzado "Google Calendar API" (identificador
//  Calendar) en vez del servicio básico CalendarApp. Ver la nota de acceso
//  multiusuario al inicio de este archivo: es lo que permite que el script
//  funcione igual sin importar qué cuenta lo ejecute, siempre que tenga
//  permiso de "Hacer cambios en eventos" sobre el calendario. Recuerda
//  habilitar el servicio en Extensiones > Apps Script > Servicios > Google
//  Calendar API antes de ejecutar el script.
// ----------------------------------------------------------------------------

/**
 * Trae, con paginación, todos los eventos del calendario cuyo horario cae
 * dentro de [timeMin, timeMax).
 */
function listarEventosEnRango(calendarId, timeMin, timeMax) {
  const eventos = [];
  let pageToken = null;

  do {
    const respuesta = Calendar.Events.list(calendarId, {
      timeMin: timeMin.toISOString(),
      timeMax: timeMax.toISOString(),
      singleEvents: true,
      maxResults: 2500,
      pageToken: pageToken || undefined,
    });

    (respuesta.items || []).forEach((evento) => eventos.push(evento));
    pageToken = respuesta.nextPageToken || null;
  } while (pageToken);

  return eventos;
}

/**
 * Incorpora un evento a los dos índices en memoria: por ID (búsqueda directa
 * cuando la fila ya tiene un ID Evento guardado) y por día (respaldo para
 * filas que todavía no tienen ID guardado, buscando por título y fecha).
 */
function registrarEventoEnIndices(evento, eventosPorId, eventosPorDia) {
  eventosPorId.set(evento.id, evento);

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
  if (evento.start && evento.start.date && evento.end && evento.end.date) {
    return {
      inicio: parsearFechaISOaDate(evento.start.date),
      fin: parsearFechaISOaDate(evento.end.date),
    };
  }

  if (evento.start && evento.start.dateTime && evento.end && evento.end.dateTime) {
    const inicio = new Date(evento.start.dateTime);
    inicio.setHours(0, 0, 0, 0);
    const fin = new Date(evento.end.dateTime);
    fin.setHours(0, 0, 0, 0);
    fin.setDate(fin.getDate() + 1);
    return { inicio, fin };
  }

  return null;
}

/** Convierte una fecha 'aaaa-mm-dd' (formato de la Calendar API) a un Date local. */
function parsearFechaISOaDate(fechaTexto) {
  const partes = fechaTexto.split('-');
  return crearFechaLocal(partes[0], partes[1], partes[2]);
}

/**
 * Recupera el evento ya asociado a la fila. Primero por el ID guardado en la
 * columna "ID Evento" (mecanismo principal, inmune a cambios de título o
 * fecha), consultando primero el índice en memoria y, solo si no aparece ahí,
 * a la API directamente. Si la fila todavía no tiene ID guardado (por
 * ejemplo, filas creadas con una versión anterior del script), se recurre a
 * la búsqueda por título y fecha para no duplicar eventos ya existentes.
 *
 * @return {?Object} el recurso del evento existente, o null si no hay ninguno.
 */
function obtenerEventoExistente(calendarId, eventosPorId, eventosPorDia, idGuardado, titulo, fechaInfo) {
  if (idGuardado) {
    if (eventosPorId.has(idGuardado)) return eventosPorId.get(idGuardado);
    try {
      return Calendar.Events.get(calendarId, idGuardado);
    } catch (err) {
      // El evento guardado ya no existe en el calendario; se buscará por
      // título/fecha o se creará uno nuevo.
    }
  }

  const candidatos = eventosPorDia.get(formatearClaveDia(fechaInfo.inicio)) || [];
  return candidatos.find((ev) => (ev.summary || '').trim() === titulo) || null;
}

/**
 * Crea el evento de día único o de rango según corresponda, aplicando color
 * y descripción.
 */
function crearEvento(calendarId, titulo, fechaInfo, descripcion, color) {
  const finInclusive = fechaInfo.esRango ? fechaInfo.finInclusive : fechaInfo.inicio;
  const finExclusivo = sumarDias(finInclusive, 1);

  const recurso = {
    summary: titulo,
    description: descripcion,
    start: { date: formatearClaveDia(fechaInfo.inicio) },
    end: { date: formatearClaveDia(finExclusivo) },
  };

  if (color) {
    recurso.colorId = color;
  }

  return Calendar.Events.insert(recurso, calendarId);
}

/**
 * Actualiza un evento existente para reflejar los datos actuales de la fila
 * (título, fecha/rango, descripción y color). Se arma un objeto con
 * únicamente los campos que efectivamente cambiaron respecto al estado
 * actual del evento, y se aplica en una sola llamada patch, para no
 * reescribir el evento completo en cada sincronización cuando la fila no
 * tuvo modificaciones.
 *
 * @return {{evento: Object, huboCambios: boolean}}
 */
function actualizarEventoSiCorresponde(calendarId, evento, titulo, fechaInfo, descripcion, color) {
  const cambios = {};

  if ((evento.summary || '') !== titulo) {
    cambios.summary = titulo;
  }

  if ((evento.description || '') !== descripcion) {
    cambios.description = descripcion;
  }

  if (fechasDistintas(evento, fechaInfo)) {
    const finExclusivo = sumarDias(fechaInfo.esRango ? fechaInfo.finInclusive : fechaInfo.inicio, 1);
    cambios.start = { date: formatearClaveDia(fechaInfo.inicio) };
    cambios.end = { date: formatearClaveDia(finExclusivo) };
  }

  if (color && evento.colorId !== color) {
    cambios.colorId = color;
  }

  const huboCambios = Object.keys(cambios).length > 0;
  const eventoActualizado = huboCambios ? Calendar.Events.patch(cambios, calendarId, evento.id) : evento;

  return { evento: eventoActualizado, huboCambios };
}

/** true si la fecha/rango del evento no coincide con la que le corresponde según la planilla. */
function fechasDistintas(evento, fechaInfo) {
  if (!evento.start || !evento.start.date || !evento.end || !evento.end.date) return true;

  const finDeseadoExclusivo = sumarDias(fechaInfo.esRango ? fechaInfo.finInclusive : fechaInfo.inicio, 1);
  return (
    evento.start.date !== formatearClaveDia(fechaInfo.inicio) ||
    evento.end.date !== formatearClaveDia(finDeseadoExclusivo)
  );
}

/** Devuelve una nueva fecha sumando N días a la fecha dada (no muta la original). */
function sumarDias(fecha, dias) {
  const nueva = new Date(fecha);
  nueva.setDate(nueva.getDate() + dias);
  return nueva;
}

/** Genera una clave "aaaa-mm-dd" a partir de una fecha; se usa como índice de
 *  día y también como formato de fecha que espera la Calendar API. */
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