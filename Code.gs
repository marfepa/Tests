const WEBAPP_BASE_URL = 'https://script.google.com/a/macros/scorazon.hhdc.net/s/AKfycby-IeBgoIb557x8OqF9taEQSnRGx0KdIsae9-OR8Nmbx04EfxwQNcLsKRTS3Fak_N8/exec';

const DESTINATION_CALIFICACIONES_SPREADSHEET_ID = '1WKVottJP88lQ-XxB2SLaLJc06aB5yQYw5peI-8WLaO0';
const DESTINATION_CALIFICACIONES_SHEET = 'CalificacionesDetalladas';

const SHEETS = {
  RESPUESTAS: 'Resultados',
  INTENTOS: 'Intentos',
  QUIZZES: 'Cuestionarios',
  PREGUNTAS: 'Preguntas',
  ESTUDIANTES: 'Estudiantes',
  CALIFICACIONES: 'Calificaciones',
};

const HEADERS = {
  [SHEETS.RESPUESTAS]: [
    'Fecha',
    'Email',
    'QuizId',
    'DuracionMin',
    'TiempoEmpleadoSeg',
    'Estado',
    'PuntajeMax',
    'PuntajeObtenido',
    'RequiereRevision',
    'DetalleJSON',
    'AlumnoId',
    'NombreEstudiante',
    'CursoEvaluado',
    'Instrumento',
    'CalificacionDetalleId',
  ],
  [SHEETS.INTENTOS]: [
    'Fecha',
    'Email',
    'QuizId',
    'Estado',
    'Notas',
    'InicioISO',
    'FinISO',
    'AlumnoId',
    'NombreEstudiante',
    'CursoEvaluado',
    'Salidas',
    'PenalizacionAcumuladaMs',
  ],
  [SHEETS.QUIZZES]: [
    'QuizId',
    'Titulo',
    'DuracionMin',
    'CursoDestino',
    'EscapeAccion',
    'EscapeValor',
    'EscapeMaxSalidas',
    'Link',
  ],
  [SHEETS.PREGUNTAS]: ['QuizId', 'Numero', 'Tipo', 'Texto', 'Opciones', 'Correctas', 'Puntaje', 'Configuracion'],
  [SHEETS.ESTUDIANTES]: ['AlumnoId', 'NombreEstudiante', 'CursoEvaluado', 'Email'],
  [SHEETS.CALIFICACIONES]: [
    'IDCalificacionDetalle',
    'IDCalificacionMaestra',
    'NombreInstrumento',
    'AlumnoEvaluador',
    'NombreEstudiante',
    'CursoEvaluado',
    'NombreSituacion',
    'FechaEvaluacion',
    'NombreCriterioEvaluado',
    'NombreNivelAlcanzado',
    'PuntuacionCriterio',
    'DescripcionItemEvaluado',
    'CompletadoItem',
    'CalificacionTotalInstrumento',
    'ComentariosGenerales',
    'ComentariosGlobales',
  ],
};

const STATUS = {
  EN_CURSO: 'EN_CURSO',
  COMPLETADO: 'COMPLETADO',
  BLOQUEADO: 'BLOQUEADO',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Test')
    .addItem('Iniciar test…', 'abrirTest')
    .addItem('Constructor de cuestionarios', 'abrirConstructor')
    .addSeparator()
    .addItem('Generar enlaces de cuestionarios', 'generarEnlaces')
    .addItem('Exportar calificaciones', 'exportarCalificaciones')
    .addSeparator()
    .addItem('Desbloquear intento', 'desbloquearIntento')
    .addSeparator()
    .addItem('Importar cuestionario desde hoja', 'importarCuestionarioDesdeHoja')
    .addToUi();
}

function abrirTest() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    'Iniciar test',
    'Escribe el QuizId exactamente como aparece en la hoja "Cuestionarios".',
    ui.ButtonSet.OK_CANCEL,
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const quizId = prompt.getResponseText().trim();
  if (!quizId) {
    ui.alert('Debes indicar un QuizId.');
    return;
  }

  const quiz = getQuizConfig(quizId);
  if (!quiz) {
    ui.alert(`No se encontró configuración válida para "${quizId}".`);
    return;
  }

  const template = HtmlService.createTemplateFromFile('test');
  template.quizId = quizId;
  template.origen = 'dialog';
  const html = template.evaluate().setWidth(520).setHeight(680);
  ui.showModalDialog(html, quiz.titulo || 'Evaluación');
}

function abrirConstructor() {
  const template = HtmlService.createTemplateFromFile('builder');
  template.webAppUrl = WEBAPP_BASE_URL;
  const html = template
    .evaluate()
    .setWidth(900)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Constructor de cuestionarios');
}

function generarEnlaces() {
  const baseUrl = WEBAPP_BASE_URL;
  if (!baseUrl) {
    throw new Error('Configura la constante WEBAPP_BASE_URL con el enlace de la aplicación web.');
  }

  const sheet = ensureSheet(SHEETS.QUIZZES, HEADERS[SHEETS.QUIZZES]);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return 'No hay cuestionarios configurados.';
  }

  const links = [];
  for (let i = 1; i < data.length; i++) {
    const quizId = (data[i][0] || '').toString().trim();
    if (!quizId) {
      links.push(['']);
      continue;
    }
    const link = `${baseUrl}?quizId=${encodeURIComponent(quizId)}`;
    links.push([link]);
  }

  sheet.getRange(2, 8, links.length, 1).setValues(links);
  return 'Enlaces actualizados.';
}

function exportarCalificaciones() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    'Exportar calificaciones',
    'Introduce el QuizId cuyas calificaciones quieres exportar.',
    ui.ButtonSet.OK_CANCEL,
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const quizId = prompt.getResponseText().trim();
  if (!quizId) {
    ui.alert('Debes indicar un QuizId.');
    return;
  }

  const quiz = getQuizConfig(quizId);
  if (!quiz) {
    ui.alert(`No se encontró configuración válida para "${quizId}".`);
    return;
  }

  const resultadosSheet = ensureSheet(SHEETS.RESPUESTAS, HEADERS[SHEETS.RESPUESTAS]);
  const data = resultadosSheet.getDataRange().getValues();
  if (data.length <= 1) {
    ui.alert('No hay resultados para exportar.');
    return;
  }

  const headerMap = HEADERS[SHEETS.RESPUESTAS].reduce((acc, value, index) => {
    acc[value] = index;
    return acc;
  }, {});

  const calificacionesSS = SpreadsheetApp.openById(DESTINATION_CALIFICACIONES_SPREADSHEET_ID);
  const calSheet = calificacionesSS.getSheetByName(DESTINATION_CALIFICACIONES_SHEET);
  if (!calSheet) {
    ui.alert(`No se encontró la pestaña "${DESTINATION_CALIFICACIONES_SHEET}" en el libro de destino.`);
    return;
  }

  if (calSheet.getLastRow() === 0) {
    calSheet
      .getRange(1, 1, 1, HEADERS[SHEETS.CALIFICACIONES].length)
      .setValues([HEADERS[SHEETS.CALIFICACIONES]]);
  }

  const existingIds = new Set();
  const calLastRow = calSheet.getLastRow();
  if (calLastRow > 1) {
    const idValues = calSheet.getRange(2, 1, calLastRow - 1, 1).getValues();
    idValues.forEach(row => {
      const id = (row[0] || '').toString().trim();
      if (id) {
        existingIds.add(id);
      }
    });
  }

  const filasAInsertar = [];
  const updates = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((row[headerMap.QuizId] || '').toString().trim() !== quizId) {
      continue;
    }
    if ((row[headerMap.Estado] || '').toString().trim() !== STATUS.COMPLETADO) {
      continue;
    }

    const alumnoId = (row[headerMap.AlumnoId] || '').toString().trim();
    const alumnoNombre = (row[headerMap.NombreEstudiante] || '').toString().trim();
    const cursoEvaluado = (row[headerMap.CursoEvaluado] || '').toString().trim();
    if (!alumnoId || !alumnoNombre) {
      continue;
    }

    let detalleId = (row[headerMap.CalificacionDetalleId] || '').toString().trim();
    if (!detalleId) {
      detalleId = Utilities.getUuid();
    }

    if (existingIds.has(detalleId)) {
      continue;
    }
    existingIds.add(detalleId);

    const fecha = row[headerMap.Fecha] instanceof Date
      ? row[headerMap.Fecha]
      : new Date(row[headerMap.Fecha] || new Date());
    const puntaje = Number(row[headerMap.PuntajeObtenido]) || 0;
    const instrumento = (row[headerMap.Instrumento] || quiz.titulo || quizId).toString();

    filasAInsertar.push([
      detalleId,
      quizId,
      instrumento,
      '',
      alumnoNombre,
      cursoEvaluado,
      '',
      fecha,
      '',
      '',
      '',
      '',
      '',
      puntaje,
      '',
      '',
    ]);

    updates.push({ rowIndex: i + 1, detalleId });
  }

  if (!filasAInsertar.length) {
    ui.alert('No se encontraron calificaciones pendientes de exportar para ese QuizId.');
    return;
  }

  const startRow = calSheet.getLastRow() + 1;
  calSheet
    .getRange(startRow, 1, filasAInsertar.length, HEADERS[SHEETS.CALIFICACIONES].length)
    .setValues(filasAInsertar);

  const colDetalleId = headerMap.CalificacionDetalleId + 1;
  updates.forEach(update => {
    resultadosSheet.getRange(update.rowIndex, colDetalleId).setValue(update.detalleId);
  });

  ui.alert(`Se exportaron ${filasAInsertar.length} calificaciones del cuestionario ${quizId}.`);
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('test');
  template.quizId = (e && e.parameter && e.parameter.quizId) || '';
  template.origen = 'webapp';
  return template.evaluate().setTitle('Evaluación');
}

function obtenerEstado(quizId, clientToken) {
  if (!quizId) {
    return {
      email: '',
      bloqueado: true,
      motivo: 'No se indicó un ID de cuestionario.',
      estudiantes: [],
      alumnoId: '',
      alumnoNombre: '',
      alumnoCurso: '',
    };
  }

  const quiz = getQuizConfig(quizId);
  if (!quiz) {
    return {
      email: '',
      bloqueado: true,
      motivo: `El cuestionario "${quizId}" no está configurado o no tiene preguntas.`,
      estudiantes: [],
      alumnoId: '',
      alumnoNombre: '',
      alumnoCurso: '',
    };
  }

  const estudiantes = getStudents(quiz.cursoDestino);

  const email = Session.getActiveUser().getEmail();
  if (!email) {
    return {
      email: '',
      bloqueado: true,
      motivo: 'No se pudo obtener tu correo. Inicia sesión con la cuenta del dominio.',
      estudiantes,
      alumnoId: '',
      alumnoNombre: '',
      alumnoCurso: '',
    };
  }

  if (quiz.cursoDestino && estudiantes.length === 0) {
    return {
      email,
      bloqueado: true,
      motivo: `No hay estudiantes disponibles para el curso "${quiz.cursoDestino}". Revisa la pestaña "Estudiantes".`,
      estudiantes,
      alumnoId: '',
      alumnoNombre: '',
      alumnoCurso: '',
    };
  }

  const attemptsSheet = ensureSheet(SHEETS.INTENTOS, HEADERS[SHEETS.INTENTOS]);
  let match = findAttemptRow(attemptsSheet, email, quizId);
  const now = new Date();
  const limitMs = quiz.duracionMin > 0 ? quiz.duracionMin * 60 * 1000 : null;
  const normalizedToken = clientToken ? clientToken.toString().trim() : '';

  if (!match) {
    const startIso = now.toISOString();
    const generatedToken = normalizedToken || Utilities.getUuid();
    const noteValue = composeAttemptNoteCell({
      mensaje: 'Intento inicial',
      token: generatedToken,
    });
    attemptsSheet.appendRow([
      new Date(),
      email,
      quizId,
      STATUS.EN_CURSO,
      noteValue,
      startIso,
      '',
      '',
      '',
      '',
      0, // Salidas
      0, // PenalizacionAcumuladaMs
    ]);
    return {
      email,
      bloqueado: false,
      quiz,
      tiempoRestanteMs: limitMs,
      estudiantes,
      alumnoId: '',
      alumnoNombre: '',
      alumnoCurso: '',
      sesionToken: generatedToken,
    };
  }

  const attemptToken = match.note.token || '';
  const noteMessage = match.note.mensaje || '';

  if (match.status === STATUS.COMPLETADO) {
    return {
      email,
      bloqueado: true,
      motivo: 'Ya completaste este cuestionario.',
      estudiantes,
      alumnoId: match.alumnoId || '',
      alumnoNombre: match.alumnoNombre || '',
      alumnoCurso: match.alumnoCurso || '',
      sesionToken: '',
    };
  }

  if (match.status === STATUS.BLOQUEADO) {
    return {
      email,
      bloqueado: true,
      motivo: match.nota || 'Tu intento anterior quedó bloqueado.',
      estudiantes,
      alumnoId: match.alumnoId || '',
      alumnoNombre: match.alumnoNombre || '',
      alumnoCurso: match.alumnoCurso || '',
      sesionToken: '',
    };
  }

  if (attemptToken) {
    if (!normalizedToken) {
      lockAttempt(
        attemptsSheet,
        match.row,
        STATUS.BLOQUEADO,
        'Se detectó otro acceso con esta cuenta.',
        now,
        {
          token: attemptToken,
          bloqueo: 'DUPLICADO',
        },
      );
      return {
        email,
        bloqueado: true,
        motivo: 'Se detectó otro acceso con esta cuenta. El intento ha sido bloqueado.',
        estudiantes,
        alumnoId: match.alumnoId || '',
        alumnoNombre: match.alumnoNombre || '',
        alumnoCurso: match.alumnoCurso || '',
        sesionToken: '',
      };
    }

    if (attemptToken !== normalizedToken) {
      lockAttempt(
        attemptsSheet,
        match.row,
        STATUS.BLOQUEADO,
        'Se detectó otro acceso con esta cuenta.',
        now,
        {
          token: attemptToken,
          bloqueo: 'DUPLICADO',
        },
      );
      return {
        email,
        bloqueado: true,
        motivo: 'Se detectó otro acceso con esta cuenta. El intento ha sido bloqueado.',
        estudiantes,
        alumnoId: match.alumnoId || '',
        alumnoNombre: match.alumnoNombre || '',
        alumnoCurso: match.alumnoCurso || '',
        sesionToken: '',
      };
    }
  }

  let activeToken = attemptToken;
  if (!activeToken) {
    activeToken = normalizedToken || Utilities.getUuid();
    const updatedNote = composeAttemptNoteCell({
      mensaje: noteMessage || 'Intento en curso',
      token: activeToken,
    });
    attemptsSheet.getRange(match.row, 5).setValue(updatedNote);
    match.note = parseAttemptNoteCell(updatedNote);
    match.nota = match.note.mensaje;
  }

  let startTime = match.inicio ? new Date(match.inicio) : now;
  if (!match.inicio) {
    attemptsSheet.getRange(match.row, 6).setValue(now.toISOString());
    startTime = now;
  }

  if (limitMs) {
    const transcurrido = now.getTime() - startTime.getTime();
    const restante = limitMs - transcurrido;
    if (restante <= 0) {
      lockAttempt(
        attemptsSheet,
        match.row,
        STATUS.BLOQUEADO,
        'Tiempo agotado automáticamente',
        now,
        {
          bloqueo: 'TIEMPO',
        },
      );
      return {
        email,
        bloqueado: true,
        motivo: 'Se agotó el tiempo para este cuestionario.',
        estudiantes,
        alumnoId: match.alumnoId || '',
        alumnoNombre: match.alumnoNombre || '',
        alumnoCurso: match.alumnoCurso || '',
        sesionToken: '',
      };
    }
    return {
      email,
      bloqueado: false,
      quiz,
      tiempoRestanteMs: restante,
      estudiantes,
      alumnoId: match.alumnoId || '',
      alumnoNombre: match.alumnoNombre || '',
      alumnoCurso: match.alumnoCurso || '',
      sesionToken: activeToken,
    };
  }

  return {
    email,
    bloqueado: false,
    quiz,
    tiempoRestanteMs: null,
    estudiantes,
    alumnoId: match.alumnoId || '',
    alumnoNombre: match.alumnoNombre || '',
    alumnoCurso: match.alumnoCurso || '',
    sesionToken: activeToken,
  };
}

function guardarRespuestas(quizId, respuestas, tiempoEmpleadoMs, alumnoId, clientToken) {
  if (!quizId) {
    throw new Error('No se recibió el ID del cuestionario.');
  }

  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('No se pudo identificar al usuario.');
  }

  if (!alumnoId) {
    throw new Error('Selecciona un estudiante antes de enviar.');
  }

  const quiz = getQuizConfig(quizId);
  if (!quiz) {
    throw new Error(`El cuestionario "${quizId}" no está configurado.`);
  }

  const estudiante = getStudentById(alumnoId);
  if (!estudiante) {
    throw new Error('No se encontró el estudiante seleccionado. Verifica la pestaña "Estudiantes".');
  }

  const filtrosCurso = parseCourseFilter(quiz.cursoDestino);
  if (filtrosCurso.length) {
    const cursoNormalizado = normalizarTextoCurso(estudiante.curso);
    const coincideCurso = filtrosCurso.some(filtro => filtro === cursoNormalizado);
    if (!coincideCurso) {
      throw new Error('El estudiante seleccionado no pertenece al curso permitido para este cuestionario.');
    }
  }

  const attemptsSheet = ensureSheet(SHEETS.INTENTOS, HEADERS[SHEETS.INTENTOS]);
  const match = findAttemptRow(attemptsSheet, email, quizId);
  if (!match || match.status !== STATUS.EN_CURSO) {
    throw new Error('Tu intento no está disponible o ya fue cerrado.');
  }

  const normalizedToken = clientToken ? clientToken.toString().trim() : '';
  if (!normalizedToken) {
    throw new Error('No se pudo validar la sesión activa. Recarga el formulario.');
  }

  const now = new Date();
  const attemptToken = match.note.token || '';

  if (!attemptToken) {
    const updatedNote = composeAttemptNoteCell({
      mensaje: match.nota || 'Intento en curso',
      token: normalizedToken,
    });
    attemptsSheet.getRange(match.row, 5).setValue(updatedNote);
    match.note = parseAttemptNoteCell(updatedNote);
    match.nota = match.note.mensaje;
  } else if (attemptToken !== normalizedToken) {
    lockAttempt(
      attemptsSheet,
      match.row,
      STATUS.BLOQUEADO,
      'Se detectó otro acceso con esta cuenta al enviar.',
      now,
      {
        token: attemptToken,
        bloqueo: 'DUPLICADO',
      },
    );
    throw new Error('El intento fue bloqueado por detectarse otro acceso simultáneo con esta cuenta.');
  }

  const startTime = match.inicio ? new Date(match.inicio) : now;
  const limitMs = quiz.duracionMin > 0 ? quiz.duracionMin * 60 * 1000 : null;
  const empleado = typeof tiempoEmpleadoMs === 'number' && tiempoEmpleadoMs > 0
    ? tiempoEmpleadoMs
    : now.getTime() - startTime.getTime();

  if (limitMs && now.getTime() - startTime.getTime() > limitMs + 2000) {
    lockAttempt(
      attemptsSheet,
      match.row,
      STATUS.BLOQUEADO,
      'Tiempo agotado al enviar respuestas',
      now,
      {
        bloqueo: 'TIEMPO',
      },
    );
    throw new Error('Las respuestas llegaron después de agotar el tiempo.');
  }

  const questionMap = new Map();
  quiz.preguntas.forEach(p => {
    questionMap.set(String(p.numero), p);
  });

  let puntajeMax = 0;
  let puntajeObtenido = 0;
  let requiereRevision = false;
  const detalles = [];

  (respuestas || []).forEach(respuesta => {
    const numeroClave = String(respuesta.numero);
    const definicion = questionMap.get(numeroClave);
    if (!definicion) {
      detalles.push({
        numero: respuesta.numero,
        tipo: respuesta.tipo,
        respuesta: respuesta.respuesta,
        encontrado: false,
      });
      return;
    }

    const evaluacion = evaluarRespuesta(definicion, respuesta.respuesta);
    puntajeMax += definicion.puntaje || 0;
    puntajeObtenido += evaluacion.puntaje || 0;
    requiereRevision = requiereRevision || evaluacion.requiereRevision;

    detalles.push({
      numero: definicion.numero,
      tipo: definicion.tipo,
      respuesta: respuesta.respuesta,
      opciones: definicion.opciones,
      correctas: definicion.correctas,
      puntajePregunta: definicion.puntaje,
      ...evaluacion,
    });
  });

  const resultsSheet = ensureSheet(SHEETS.RESPUESTAS, HEADERS[SHEETS.RESPUESTAS]);
  resultsSheet.appendRow([
    new Date(),
    email,
    quizId,
    quiz.duracionMin,
    Math.round(empleado / 1000),
    STATUS.COMPLETADO,
    puntajeMax,
    puntajeObtenido,
    requiereRevision ? 'SI' : 'NO',
    JSON.stringify(detalles),
    estudiante.id,
    estudiante.nombre,
    estudiante.curso,
    quiz.titulo || quizId,
    '',
  ]);

  attemptsSheet.getRange(match.row, 8, 1, 3).setValues([[estudiante.id, estudiante.nombre, estudiante.curso]]);
  lockAttempt(
    attemptsSheet,
    match.row,
    STATUS.COMPLETADO,
    requiereRevision ? 'Requiere revisión manual' : 'Respuestas entregadas',
    now,
    {
      bloqueo: 'FINALIZADO',
    },
  );
}

function registrarEscape(quizId, motivo, clientToken) {
  if (!quizId) {
    return;
  }

  const email = Session.getActiveUser().getEmail();
  if (!email) {
    return;
  }

  const attemptsSheet = ensureSheet(SHEETS.INTENTOS, HEADERS[SHEETS.INTENTOS]);
  const match = findAttemptRow(attemptsSheet, email, quizId);
  const nota = motivo || 'Intento bloqueado';
  const now = new Date();
  const normalizedToken = clientToken ? clientToken.toString().trim() : '';

  if (!match) {
    const noteValue = composeAttemptNoteCell({
      mensaje: nota,
      token: normalizedToken,
      bloqueo: 'ESCAPE',
    });
    attemptsSheet.appendRow([
      new Date(),
      email,
      quizId,
      STATUS.BLOQUEADO,
      noteValue,
      now.toISOString(),
      now.toISOString(),
      '',
      '',
      '',
    ]);
    return;
  }

  if (match.status === STATUS.EN_CURSO) {
    const attemptToken = match.note.token || '';
    if (attemptToken && normalizedToken && attemptToken !== normalizedToken) {
      lockAttempt(
        attemptsSheet,
        match.row,
        STATUS.BLOQUEADO,
        'Se detectó otro acceso con esta cuenta.',
        now,
        {
          token: attemptToken,
          bloqueo: 'DUPLICADO',
        },
      );
      return;
    }

    const tokenToStore = attemptToken || normalizedToken;
    lockAttempt(
      attemptsSheet,
      match.row,
      STATUS.BLOQUEADO,
      nota,
      now,
      {
        token: tokenToStore,
        bloqueo: 'ESCAPE',
      },
    );
  }
}

function asignarEstudiante(quizId, alumnoId, clientToken) {
  if (!quizId) {
    throw new Error('Falta el identificador del cuestionario.');
  }
  if (!alumnoId) {
    throw new Error('Selecciona un estudiante.');
  }

  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('No se pudo identificar al usuario.');
  }

  const quiz = getQuizConfig(quizId);
  if (!quiz) {
    throw new Error('No se encontró la configuración del cuestionario.');
  }

  const estudiante = getStudentById(alumnoId);
  if (!estudiante) {
    throw new Error('No se encontró el estudiante en la hoja "Estudiantes".');
  }

  const filtrosCurso = parseCourseFilter(quiz.cursoDestino);
  if (filtrosCurso.length) {
    const cursoNormalizado = normalizarTextoCurso(estudiante.curso);
    const coincideCurso = filtrosCurso.some(filtro => filtro === cursoNormalizado);
    if (!coincideCurso) {
      throw new Error('Este estudiante no pertenece al curso asignado para el cuestionario.');
    }
  }

  const attemptsSheet = ensureSheet(SHEETS.INTENTOS, HEADERS[SHEETS.INTENTOS]);
  let match = findAttemptRow(attemptsSheet, email, quizId);

  if (!match) {
    throw new Error('No se encontró un intento activo. Recarga el formulario.');
  }

  if (match.status !== STATUS.EN_CURSO) {
    throw new Error('El intento ya no está disponible para editar.');
  }

  const normalizedToken = clientToken ? clientToken.toString().trim() : '';
  const attemptToken = match.note.token || '';

  if (!attemptToken && !normalizedToken) {
    throw new Error('No se pudo validar la sesión activa. Recarga el formulario.');
  }

  if (attemptToken && normalizedToken && attemptToken !== normalizedToken) {
    throw new Error('No se pudo validar la sesión activa del intento.');
  }

  if (!attemptToken && normalizedToken) {
    const updatedNote = composeAttemptNoteCell({
      mensaje: match.nota || 'Intento en curso',
      token: normalizedToken,
    });
    attemptsSheet.getRange(match.row, 5).setValue(updatedNote);
    match.note = parseAttemptNoteCell(updatedNote);
    match.nota = match.note.mensaje;
  }

  attemptsSheet.getRange(match.row, 8, 1, 3).setValues([[
    estudiante.id,
    estudiante.nombre,
    estudiante.curso,
  ]]);

  return estudiante;
}

function listarCuestionarios() {
  const sheet = ensureSheet(SHEETS.QUIZZES, HEADERS[SHEETS.QUIZZES]);
  const data = sheet.getDataRange().getValues();
  const quizzes = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const quizId = (row[0] || '').toString().trim();
    if (!quizId) {
      continue;
    }

    const cursoDestinoRaw = (row[3] || '').toString().trim();
    const cursoDestino = /^https?:\/\//i.test(cursoDestinoRaw) ? '' : cursoDestinoRaw;
    const link = (row[7] || cursoDestinoRaw || '').toString();

    quizzes.push({
      quizId,
      titulo: (row[1] || '').toString(),
      duracionMin: Number(row[2]) || 0,
      cursoDestino,
      link,
      escapeConfig: buildEscapeConfig(row[4], row[5], row[6]),
    });
  }

  quizzes.sort((a, b) => a.quizId.localeCompare(b.quizId));
  return quizzes;
}

function buildEscapeConfig(accionRaw, valorRaw, maxSalidasRaw) {
  const accion = (accionRaw || 'PENALIZACION').toString().trim().toUpperCase();
  const valor = (valorRaw || '').toString().trim();
  let maxSalidas = Number(maxSalidasRaw);
  if (!isFinite(maxSalidas) || maxSalidas <= 0) {
    maxSalidas = accion === 'PENALIZACION' ? 2 : 1;
  }

  return {
    accion,
    valor,
    maxSalidas,
  };
}

function obtenerQuizCompleto(quizId) {
  if (!quizId) {
    return null;
  }
  const meta = getQuizMeta(quizId);
  if (!meta) {
    return null;
  }
  const preguntas = getQuizQuestions(quizId);
  return {
    quizId,
    titulo: meta.titulo,
    duracionMin: meta.duracionMin,
    cursoDestino: meta.cursoDestino || '',
    escapeConfig: meta.escapeConfig,
    preguntas,
  };
}

function guardarQuizCompleto(payload) {
  if (!payload || !payload.quizId) {
    throw new Error('Falta el identificador del cuestionario.');
  }

  const quizId = payload.quizId.toString().trim();
  const titulo = (payload.titulo || '').toString().trim();
  const duracionMin = Number(payload.duracionMin) || 0;
  const cursoDestino = (payload.cursoDestino || '').toString().trim();
  const escapeConfigInput = payload.escapeConfig && typeof payload.escapeConfig === 'object' ? payload.escapeConfig : {};
  const escapeConfig = buildEscapeConfig(escapeConfigInput.accion, escapeConfigInput.valor, escapeConfigInput.maxSalidas);
  const preguntas = Array.isArray(payload.preguntas) ? payload.preguntas : [];

  if (!quizId) {
    throw new Error('El QuizId es obligatorio.');
  }
  if (!/^[-_a-zA-Z0-9]+$/.test(quizId)) {
    throw new Error('El QuizId solo puede contener letras, números, guiones y guion bajo.');
  }
  if (!titulo) {
    throw new Error('El título es obligatorio.');
  }
  if (preguntas.length === 0) {
    throw new Error('Añade al menos una pregunta antes de guardar.');
  }

  const cleanedQuestions = preguntas.map((pregunta, index) => sanitizeQuestion(pregunta, index));

  const quizSheet = ensureSheet(SHEETS.QUIZZES, HEADERS[SHEETS.QUIZZES]);
  const preguntasSheet = ensureSheet(SHEETS.PREGUNTAS, HEADERS[SHEETS.PREGUNTAS]);

  upsertQuizRow(quizSheet, quizId, titulo, duracionMin, cursoDestino, escapeConfig);
  replaceQuizQuestions(preguntasSheet, quizId, cleanedQuestions);

  return {
    ok: true,
    mensaje: 'Cuestionario guardado correctamente.',
    preguntas: cleanedQuestions.length,
    cursoDestino,
    escapeConfig,
  };
}

function eliminarQuiz(quizId) {
  if (!quizId) {
    throw new Error('Indica el cuestionario a eliminar.');
  }
  const sheetQuiz = ensureSheet(SHEETS.QUIZZES, HEADERS[SHEETS.QUIZZES]);
  const sheetPreguntas = ensureSheet(SHEETS.PREGUNTAS, HEADERS[SHEETS.PREGUNTAS]);

  const dataQuiz = sheetQuiz.getDataRange().getValues();
  let removed = false;
  for (let i = dataQuiz.length - 1; i >= 1; i--) {
    const id = (dataQuiz[i][0] || '').toString().trim();
    if (id === quizId) {
      sheetQuiz.deleteRow(i + 1);
      removed = true;
    }
  }

  const dataPreguntas = sheetPreguntas.getDataRange().getValues();
  for (let i = dataPreguntas.length - 1; i >= 1; i--) {
    const id = (dataPreguntas[i][0] || '').toString().trim();
    if (id === quizId) {
      sheetPreguntas.deleteRow(i + 1);
    }
  }

  if (!removed) {
    throw new Error('No se encontró el cuestionario indicado.');
  }

  return {
    ok: true,
    mensaje: 'Cuestionario eliminado. Los intentos y resultados históricos se mantienen.',
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getStudents(cursoFiltro) {
  const sheet = ensureSheet(SHEETS.ESTUDIANTES, HEADERS[SHEETS.ESTUDIANTES]);
  const data = sheet.getDataRange().getValues();
  const estudiantes = [];

  const filtros = parseCourseFilter(cursoFiltro);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = (row[0] || '').toString().trim();
    const nombre = (row[1] || '').toString().trim();
    if (!id || !nombre) {
      continue;
    }
    const curso = (row[2] || '').toString().trim();
    if (filtros.length) {
      const cursoNormalizado = normalizarTextoCurso(curso);
      const coincide = filtros.some(filtro => filtro === cursoNormalizado);
      if (!coincide) {
        continue;
      }
    }

    estudiantes.push({
      id,
      nombre,
      curso,
      email: (row[3] || '').toString().trim(),
    });
  }

  estudiantes.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));
  return estudiantes;
}

function parseCourseFilter(cursoFiltro) {
  if (!cursoFiltro) {
    return [];
  }
  return cursoFiltro
    .split(/[,;\n]/)
    .map(item => item.trim())
    .filter(Boolean)
    .map(normalizarTextoCurso);
}

function normalizarTextoCurso(valor) {
  return valor
    .toString()
    .trim()
    .toLowerCase();
}

function getStudentById(alumnoId) {
  if (!alumnoId) {
    return null;
  }
  const estudiantes = getStudents();
  for (let i = 0; i < estudiantes.length; i++) {
    if (estudiantes[i].id === alumnoId) {
      return estudiantes[i];
    }
  }
  return null;
}

function getQuizConfig(quizId) {
  const meta = getQuizMeta(quizId);
  if (!meta) {
    return null;
  }
  const preguntas = getQuizQuestions(quizId);
  if (!preguntas.length) {
    return null;
  }

  const puntajeTotal = preguntas.reduce((acc, pregunta) => acc + (pregunta.puntaje || 0), 0);
  return {
    id: quizId,
    titulo: meta.titulo,
    duracionMin: meta.duracionMin,
    cursoDestino: meta.cursoDestino || '',
    preguntas,
    puntajeTotal,
    escapeConfig: meta.escapeConfig,
  };
}

function getQuizMeta(quizId) {
  const sheet = ensureSheet(SHEETS.QUIZZES, HEADERS[SHEETS.QUIZZES]);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowQuizId = (data[i][0] || '').toString().trim();
    if (rowQuizId && rowQuizId === quizId) {
      const cursoDestinoRaw = (data[i][3] || '').toString().trim();
      const cursoDestino = /^https?:\/\//i.test(cursoDestinoRaw) ? '' : cursoDestinoRaw;
      const escapeAccion = (data[i][4] || '').toString().trim();
      const escapeValor = (data[i][5] || '').toString().trim();
      const escapeMax = data[i][6];

      return {
        titulo: (data[i][1] || '').toString(),
        duracionMin: Number(data[i][2]) || 0,
        cursoDestino,
        escapeConfig: buildEscapeConfig(escapeAccion, escapeValor, escapeMax),
        link: (data[i][7] || '').toString(),
      };
    }
  }
  return null;
}

function getQuizQuestions(quizId) {
  const sheet = ensureSheet(SHEETS.PREGUNTAS, HEADERS[SHEETS.PREGUNTAS]);
  const data = sheet.getDataRange().getValues();
  const preguntas = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowQuizId = (row[0] || '').toString().trim();
    if (rowQuizId !== quizId) {
      continue;
    }
    const numero = Number(row[1]);
    const tipo = (row[2] || 'texto').toString().trim().toLowerCase();
    const texto = (row[3] || '').toString().trim();
    const opciones = (row[4] || '')
      .toString()
      .split('|')
      .map(s => s.trim())
      .filter(Boolean);
    const correctas = (row[5] || '')
      .toString()
      .split('|')
      .map(s => s.trim())
      .filter(Boolean);
    const puntaje = Number(row[6]) || 0;
    let config = {};
    if (row.length >= 8) {
      const rawConfig = row[7];
      if (rawConfig && rawConfig.toString().trim()) {
        try {
          const parsed = JSON.parse(rawConfig);
          if (parsed && typeof parsed === 'object') {
            config = parsed;
          }
        } catch (err) {
          config = {};
        }
      }
    }

    if (!texto) {
      continue;
    }

    preguntas.push({ numero, tipo, texto, opciones, correctas, puntaje, config });
  }

  preguntas.sort((a, b) => {
    const aNum = isFinite(a.numero) ? Number(a.numero) : 0;
    const bNum = isFinite(b.numero) ? Number(b.numero) : 0;
    return aNum - bNum;
  });

  return preguntas;
}

function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  ensureHeaders(sheet, headers);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  const neededLength = headers.length;
  const lastCol = Math.max(sheet.getLastColumn(), neededLength);
  const current = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  let update = false;
  const newRow = [];
  for (let i = 0; i < neededLength; i++) {
    const desired = headers[i];
    const existing = (current[i] || '').toString().trim();
    if (existing !== desired) {
      update = true;
    }
    newRow[i] = desired;
  }

  if (update || sheet.getLastColumn() < neededLength) {
    sheet.getRange(1, 1, 1, neededLength).setValues([newRow]);
  }
}

function findAttemptRow(sheet, email, quizId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowEmail = (row[1] || '').toString().trim();
    const rowQuizId = (row[2] || '').toString().trim();
    if (rowEmail === email && rowQuizId === quizId) {
      const noteRaw = row[4];
      const noteInfo = parseAttemptNoteCell(noteRaw);
      return {
        row: i + 2,
        status: (row[3] || '').toString().trim(),
        nota: noteInfo.mensaje || (noteRaw || '').toString(),
        note: noteInfo,
        inicio: row[5] || '',
        fin: row[6] || '',
        alumnoId: (row[7] || '').toString(),
        alumnoNombre: (row[8] || '').toString(),
        alumnoCurso: (row[9] || '').toString(),
      };
    }
  }
  return null;
}

function lockAttempt(sheet, row, status, note, finDate, options) {
  sheet.getRange(row, 4).setValue(status);
  const config = options || {};
  const noteValue = composeAttemptNoteCell({
    mensaje: note,
    token: config.token || '',
    bloqueo: config.bloqueo || '',
  });
  sheet.getRange(row, 5).setValue(noteValue);
  if (finDate) {
    sheet.getRange(row, 7).setValue(finDate.toISOString());
  }
}

function composeAttemptNoteCell(data) {
  const mensaje = data && data.mensaje ? data.mensaje : '';
  const token = data && data.token ? data.token : '';
  const bloqueo = data && data.bloqueo ? data.bloqueo : '';

  const hasToken = Boolean(token);
  const hasBloqueo = Boolean(bloqueo);

  if (!hasToken && !hasBloqueo) {
    return mensaje || '';
  }

  const payload = {};
  if (mensaje) {
    payload.mensaje = mensaje;
  }
  if (hasToken) {
    payload.token = token;
  }
  if (hasBloqueo) {
    payload.bloqueo = bloqueo;
  }

  return JSON.stringify(payload);
}

function parseAttemptNoteCell(value) {
  const empty = {
    mensaje: '',
    token: '',
    bloqueo: '',
  };

  if (value === null || value === undefined) {
    return empty;
  }

  const text = value.toString().trim();
  if (!text) {
    return empty;
  }

  if (text.startsWith('{')) {
    try {
      const parsed = JSON.parse(text);
      if (parsed && typeof parsed === 'object') {
        return {
          mensaje: parsed.mensaje ? parsed.mensaje.toString() : '',
          token: parsed.token ? parsed.token.toString() : '',
          bloqueo: parsed.bloqueo ? parsed.bloqueo.toString() : '',
        };
      }
    } catch (err) {
      return {
        mensaje: text,
        token: '',
        bloqueo: '',
      };
    }
  }

  return {
    mensaje: text,
    token: '',
    bloqueo: '',
  };
}

function sanitizeQuestion(question, index) {
  const tipo = (question.tipo || 'texto').toString().toLowerCase();
  const texto = (question.texto || '').toString().trim();
  const puntaje = Number(question.puntaje);

  if (!texto) {
    throw new Error(`La pregunta ${index + 1} no tiene enunciado.`);
  }

  const sanitized = {
    numero: index + 1,
    tipo,
    texto,
    opciones: [],
    correctas: [],
    puntaje: isFinite(puntaje) && puntaje >= 0 ? puntaje : 1,
    config: {},
  };

  if (tipo === 'radio' || tipo === 'opcion' || tipo === 'checkbox') {
    const opciones = Array.isArray(question.opciones) ? question.opciones : [];
    const filtradas = opciones
      .map(op => sanitizeCellText(op))
      .filter(Boolean);
    if (filtradas.length < 2) {
      throw new Error(`La pregunta ${index + 1} necesita al menos dos opciones.`);
    }
    sanitized.opciones = filtradas;

    const correctas = Array.isArray(question.correctas) ? question.correctas : [];
    const correctasText = correctas
      .map(value => sanitizeCellText(value))
      .filter(Boolean);

    if (tipo === 'radio' || tipo === 'opcion') {
      if (correctasText.length !== 1) {
        throw new Error(`Selecciona una única respuesta correcta en la pregunta ${index + 1}.`);
      }
      if (!filtradas.includes(correctasText[0])) {
        throw new Error(`La respuesta correcta de la pregunta ${index + 1} no coincide con ninguna opción.`);
      }
    }

    if (tipo === 'checkbox' && correctasText.length === 0) {
      throw new Error(`Indica al menos una respuesta correcta para la pregunta ${index + 1}.`);
    }

    sanitized.correctas = correctasText;
  } else if (tipo === 'texto') {
    const aceptadas = Array.isArray(question.correctas) ? question.correctas : [];
    const filtradas = aceptadas
      .map(value => sanitizeCellText(value))
      .filter(Boolean);
    if (!filtradas.length) {
      throw new Error(`Añade al menos una respuesta aceptada en la pregunta ${index + 1}.`);
    }
    sanitized.correctas = filtradas;
  } else if (tipo === 'parrafo') {
    sanitized.correctas = [];
  } else if (tipo === 'relacion') {
    const izquierdas = Array.isArray(question.opciones) ? question.opciones.map(value => sanitizeCellText(value)).filter(Boolean) : [];
    const derechas = Array.isArray(question.correctas) ? question.correctas.map(value => sanitizeCellText(value)).filter(Boolean) : [];
    if (izquierdas.length < 2) {
      throw new Error(`La pregunta ${index + 1} requiere al menos dos elementos a relacionar.`);
    }
    if (izquierdas.length !== derechas.length) {
      throw new Error(`El número de elementos de la relación debe coincidir en la pregunta ${index + 1}.`);
    }
    sanitized.opciones = izquierdas;
    sanitized.correctas = derechas;
    const relaciones = izquierdas.map((izquierda, idx) => ({ izquierda, derecha: derechas[idx] }));
    sanitized.config = { relaciones };
  } else if (tipo === 'escala') {
    const cfg = question.config && typeof question.config === 'object' ? question.config : {};
    const min = Number(cfg.min);
    const max = Number(cfg.max);
    const step = cfg.step !== undefined ? Number(cfg.step) : 1;
    if (!isFinite(min) || !isFinite(max) || min >= max) {
      throw new Error(`Configura valores mínimo y máximo válidos para la escala en la pregunta ${index + 1}.`);
    }
    const escala = {
      min,
      max,
      step: isFinite(step) && step > 0 ? step : 1,
      etiquetaMin: sanitizeCellText(cfg.etiquetaMin || ''),
      etiquetaMax: sanitizeCellText(cfg.etiquetaMax || ''),
    };
    sanitized.config = { escala };
  } else {
    throw new Error(`Tipo de pregunta no soportado en la pregunta ${index + 1}.`);
  }

  return sanitized;
}

function sanitizeCellText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value
    .toString()
    .replace(/\|/g, '¦')
    .trim();
}

function upsertQuizRow(sheet, quizId, titulo, duracionMin, cursoDestino, escapeConfig) {
  const data = sheet.getDataRange().getValues();
  const accion = (escapeConfig && escapeConfig.accion) || '';
  const valor = (escapeConfig && escapeConfig.valor) || '';
  const maxSalidas = escapeConfig && Number(escapeConfig.maxSalidas) ? Number(escapeConfig.maxSalidas) : '';
  for (let i = 1; i < data.length; i++) {
    const rowId = (data[i][0] || '').toString().trim();
    if (rowId === quizId) {
      sheet.getRange(i + 1, 2).setValue(titulo);
      sheet.getRange(i + 1, 3).setValue(duracionMin);
      sheet.getRange(i + 1, 4).setValue(cursoDestino || '');
      sheet.getRange(i + 1, 5).setValue(accion);
      sheet.getRange(i + 1, 6).setValue(valor);
      sheet.getRange(i + 1, 7).setValue(maxSalidas);
      return;
    }
  }

  sheet.appendRow([quizId, titulo, duracionMin, cursoDestino || '', accion, valor, maxSalidas, '']);
}

function replaceQuizQuestions(sheet, quizId, questions) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const rowQuizId = (data[i][0] || '').toString().trim();
    if (rowQuizId === quizId) {
      sheet.deleteRow(i + 1);
    }
  }

  if (!questions.length) {
    return;
  }

  const rows = questions.map(q => [
    quizId,
    q.numero,
    q.tipo,
    q.texto,
    q.opciones.join('|'),
    q.correctas.join('|'),
    q.puntaje,
    q.config && Object.keys(q.config).length ? JSON.stringify(q.config) : '',
  ]);

  const insertRow = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(insertRow, 1, rows.length, HEADERS[SHEETS.PREGUNTAS].length).setValues(rows);
}

function evaluarRespuesta(definicion, respuesta) {
  const tipo = definicion.tipo;
  const puntaje = definicion.puntaje || 0;

  if (tipo === 'parrafo') {
    return {
      correcto: null,
      puntaje: 0,
      requiereRevision: true,
      comentario: 'Respuesta abierta, revisar manualmente.',
    };
  }

  if (tipo === 'texto') {
    const esperado = definicion.correctas.map(c => c.toLowerCase());
    const respuestaNormalizada = (respuesta || '')
      .toString()
      .trim()
      .toLowerCase();
    const esCorrecto = esperado.includes(respuestaNormalizada);
    return {
      correcto: esCorrecto,
      puntaje: esCorrecto ? puntaje : 0,
      requiereRevision: false,
    };
  }

  if (tipo === 'radio' || tipo === 'opcion') {
    const esperado = definicion.correctas[0];
    const esCorrecto = (respuesta || '').toString().trim() === esperado;
    return {
      correcto: esCorrecto,
      puntaje: esCorrecto ? puntaje : 0,
      requiereRevision: false,
    };
  }

  if (tipo === 'checkbox') {
    const respuestaLista = Array.isArray(respuesta)
      ? respuesta.map(v => v.toString().trim())
      : [respuesta].filter(Boolean).map(v => v.toString().trim());
    const esperado = (definicion.correctas || []).map(v => v.toString().trim());

    const sortFn = arr => arr.slice().sort();
    const esCorrecto = JSON.stringify(sortFn(respuestaLista)) === JSON.stringify(sortFn(esperado));

    return {
      correcto: esCorrecto,
      puntaje: esCorrecto ? puntaje : 0,
      requiereRevision: false,
    };
  }

  if (tipo === 'relacion') {
    const relaciones = definicion.config && Array.isArray(definicion.config.relaciones)
      ? definicion.config.relaciones
      : (definicion.opciones || []).map((izquierda, idx) => ({ izquierda, derecha: (definicion.correctas || [])[idx] }));
    if (!relaciones.length) {
      return {
        correcto: null,
        puntaje: 0,
        requiereRevision: true,
        comentario: 'Configuración de relación incompleta.',
      };
    }

    const respuestas = Array.isArray(respuesta) ? respuesta : [];
    let aciertos = 0;
    const mapaRespuestas = new Map();
    respuestas.forEach(item => {
      if (!item || typeof item !== 'object') { // Comprobación de seguridad
        return;
      }
      // Acepta múltiples nombres de propiedad para robustez
      const izquierda = (item.izquierda || item.left || item.key || item.clave || '').toString().trim();
      const derecha = (item.derecha || item.right || item.value || item.valor || '').toString().trim();
      if (izquierda) {
        mapaRespuestas.set(izquierda, derecha);
      }
    });

    relaciones.forEach(rel => {
      const esperadoIzquierda = (rel.izquierda || '').toString().trim();
      const esperadoDerecha = (rel.derecha || '').toString().trim();
      if (!esperadoIzquierda) return; // Ignorar si la definición es inválida
      const respuestaDerecha = mapaRespuestas.get(esperadoIzquierda) || '';
      if (respuestaDerecha && respuestaDerecha === esperadoDerecha) {
        aciertos += 1;
      }
    });

    const total = relaciones.length;
    const puntajeUnitario = total > 0 ? puntaje / total : 0;
    const puntajeObtenido = puntajeUnitario * aciertos;

    return {
      correcto: aciertos === total,
      puntaje: puntajeObtenido,
      requiereRevision: false,
      comentario: aciertos === total ? '' : `Aciertos: ${aciertos}/${total}.`,
    };
  }

  if (tipo === 'escala') {
    return {
      correcto: null,
      puntaje: 0,
      requiereRevision: true,
      comentario: 'Respuesta de escala; revisar manualmente.',
    };
  }

  return {
    correcto: null,
    puntaje: 0,
    requiereRevision: true,
    comentario: 'Tipo de pregunta no evaluado automáticamente.',
  };
}

const IMPORT_SHEET_NAME = 'Importar';

function importarCuestionarioDesdeHoja() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(IMPORT_SHEET_NAME);

  if (!importSheet) {
    ui.alert(`Para importar, primero crea una hoja llamada "${IMPORT_SHEET_NAME}" y pega ahí los datos del cuestionario.`);
    return;
  }

  const data = importSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert(`La hoja "${IMPORT_SHEET_NAME}" está vacía o solo contiene cabeceras.`);
    return;
  }

  const headers = data[0].map(h => h.toString().trim());
  const headerMap = headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});

  // Validar cabeceras necesarias
  const requiredHeaders = ['QuizId', 'TipoPregunta', 'TextoPregunta'];
  for (const h of requiredHeaders) {
    if (headerMap[h] === undefined) {
      ui.alert(`Falta la columna obligatoria "${h}" en la hoja de importación.`);
      return;
    }
  }

  const firstRow = data[1];
  const quizId = (firstRow[headerMap['QuizId']] || '').toString().trim();
  if (!quizId) {
    ui.alert('El "QuizId" es obligatorio y no puede estar vacío en la primera fila de datos.');
    return;
  }

  const result = ui.alert(
    `Se importará el cuestionario con ID: "${quizId}". Si ya existe, se sobreescribirá. ¿Continuar?`,
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  try {
    const payload = {
      quizId: quizId,
      titulo: (firstRow[headerMap['Titulo']] || quizId).toString().trim(),
      duracionMin: Number(firstRow[headerMap['DuracionMin']]) || 0,
      cursoDestino: (firstRow[headerMap['CursoDestino']] || '').toString().trim(),
      escapeConfig: {}, // Se puede extender para importar esto también
      preguntas: [],
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tipo = (row[headerMap['TipoPregunta']] || '').toString().trim().toLowerCase();
      const texto = (row[headerMap['TextoPregunta']] || '').toString().trim();

      if (!tipo || !texto) continue;

      payload.preguntas.push({
        tipo: tipo,
        texto: texto,
        opciones: (row[headerMap['Opciones']] || '').toString().split('|'),
        correctas: (row[headerMap['RespuestasCorrectas']] || '').toString().split('|'),
        puntaje: Number(row[headerMap['Puntaje']]) || 1,
        config: {} // Se puede extender para importar configs complejas (escala, etc.)
      });
    }

    const saveResult = guardarQuizCompleto(payload);
    ui.alert(`Importación completada. ${saveResult.mensaje}`);

  } catch (e) {
    ui.alert(`Error durante la importación: ${e.message}`);
  }
}

function desbloquearIntento() {
  const ui = SpreadsheetApp.getUi();
  const quizIdPrompt = ui.prompt(
    'Desbloquear intento',
    'Introduce el QuizId del cuestionario.',
    ui.ButtonSet.OK_CANCEL,
  );

  if (quizIdPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const quizId = quizIdPrompt.getResponseText().trim();
  if (!quizId) {
    ui.alert('Debes indicar un QuizId.');
    return;
  }

  const emailPrompt = ui.prompt(
    'Desbloquear intento',
    `Introduce el email del usuario para el QuizId "${quizId}".`,
    ui.ButtonSet.OK_CANCEL,
  );

  if (emailPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const email = emailPrompt.getResponseText().trim().toLowerCase();
  if (!email) {
    ui.alert('Debes indicar un email.');
    return;
  }

  const attemptsSheet = ensureSheet(SHEETS.INTENTOS, HEADERS[SHEETS.INTENTOS]);
  const match = findAttemptRow(attemptsSheet, email, quizId);

  if (!match) {
    ui.alert(`No se encontró ningún intento para "${email}" en el cuestionario "${quizId}".`);
    return;
  }

  if (match.status !== STATUS.BLOQUEADO) {
    ui.alert(`El intento de "${email}" para "${quizId}" no está bloqueado (estado actual: ${match.status}).`);
    return;
  }

  const statusColumn = HEADERS[SHEETS.INTENTOS].indexOf('Estado') + 1;
  const notesColumn = HEADERS[SHEETS.INTENTOS].indexOf('Notas') + 1;
  const inicioColumn = HEADERS[SHEETS.INTENTOS].indexOf('InicioISO') + 1;
  
  attemptsSheet.getRange(match.row, statusColumn).setValue(STATUS.EN_CURSO);
  attemptsSheet.getRange(match.row, inicioColumn).setValue(new Date().toISOString()); // Reiniciar el temporizador

  const newNote = composeAttemptNoteCell({
    mensaje: `Desbloqueado por ${Session.getActiveUser().getEmail()} el ${new Date().toLocaleString()}`,
    token: '', // Limpiar el token para evitar re-bloqueo
    bloqueo: '', // Limpiar el motivo del bloqueo
  });
  attemptsSheet.getRange(match.row, notesColumn).setValue(newNote);

  ui.alert(`El intento de "${email}" para "${quizId}" ha sido desbloqueado.`);
}
