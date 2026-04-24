const SHEETS = {
  ACTIVIDADES: 'ActividadesPOA',
  COORDINADORES: 'Coordinadores',
  REGISTROS: 'Registros',
  LISTAS: 'Listas'
};

const DRIVE_ROOT_FOLDER_ID = 'REEMPLAZAR_CON_FOLDER_ID_PRINCIPAL';
const AREAS = [
  'Ingeniería Civil',
  'CCEEyJJ',
  'Investigación',
  'Posgrado y EC',
  'BE - Proy. Social',
  'Biblioteca',
  'Supervisión Metodológica',
  'Dirección Académica',
  'Comunicación Institucional',
  'Recursos Humanos',
  'Registro Académico',
  'Ingeniería Agronómica',
  'Diseño Gráfico y Arq',
  'Gestión de Calidad',
  'TIC'
];

const HEADERS = {
  ACTIVIDADES: [
    'anio',
    'actividadId',
    'coordinacion',
    'actividad',
    'indicadorPoa',
    'ejeEne',
    'areasInvolucradas',
    'otrasAreas',
    'cuatrimestre'
  ],
  COORDINADORES: ['coordinacion', 'correo', 'activo'],
  REGISTROS: [
    'timestampRegistro',
    'registroId',
    'actividadId',
    'coordinacion',
    'correo',
    'estado',
    'fechaActividad',
    'horaInicio',
    'horaFin',
    'mes',
    'semanaMes',
    'alumnosHombres',
    'alumnasMujeres',
    'docentesHombres',
    'docentesMujeres',
    'administrativosHombres',
    'administrativasMujeres',
    'tipoActividad',
    'objetivoActividad',
    'carrerasInvolucradas',
    'areaPrincipal',
    'areasApoyo',
    'tipoProtagonista',
    'actividadNombre',
    'indicadorPoa',
    'urlsEvidencias',
    'documentoUrl'
  ]
};

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  return template
    .evaluate()
    .setTitle('Control de Actividades POA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActive();
  ensureSheetWithHeaders_(ss, SHEETS.ACTIVIDADES, HEADERS.ACTIVIDADES);
  ensureSheetWithHeaders_(ss, SHEETS.COORDINADORES, HEADERS.COORDINADORES);
  ensureSheetWithHeaders_(ss, SHEETS.REGISTROS, HEADERS.REGISTROS);
  ensureSheetWithHeaders_(ss, SHEETS.LISTAS, ['tipoActividad', 'tipoProtagonista', 'indicadorPoa', 'indicadorEstrategia', 'carreras']);
}

function getInitialData(cuatrimestreSolicitado) {
  const userEmail = Session.getActiveUser().getEmail();
  const coordinador = getCoordinatorByEmail_(userEmail);
  if (!coordinador) {
    return {
      authorized: false,
      userEmail,
      message: `El correo ${userEmail} no tiene acceso.`
    };
  }

  const cuatrimestreActual = getCurrentQuarter_();
  const cuatrimestreSeleccionado = normalizeQuarter_(cuatrimestreSolicitado, cuatrimestreActual);
  const actividadesVisibles = getVisibleActivities_(coordinador.coordinacion, cuatrimestreSeleccionado);
  const actividadesPropias = actividadesVisibles.filter(
    (a) => a.esPropietario === true
  );
  const registros = getRecordsByCoordination_(coordinador.coordinacion);
  const actividadIdsCompletadas = new Set(
    registros.filter((r) => r.estado === 'Finalizada').map((r) => r.actividadId)
  );

  const actividadesCompletadas = actividadesPropias.filter((a) =>
    actividadIdsCompletadas.has(a.actividadId)
  );
  const actividadesPendientes = actividadesPropias.filter(
    (a) => !actividadIdsCompletadas.has(a.actividadId)
  );

  const markCompletion = (activity) => ({
    ...activity,
    estaFinalizada: actividadIdsCompletadas.has(activity.actividadId)
  });

  return {
    authorized: true,
    userEmail,
    coordinacion: coordinador.coordinacion,
    cuatrimestreActual,
    cuatrimestreSeleccionado,
    actividadesPendientes: actividadesPendientes.map(markCompletion),
    actividadesCompletadas: actividadesCompletadas.map(markCompletion),
    actividadesParticipante: actividadesVisibles.filter(
      (a) => a.esPropietario === false
    ).map(markCompletion),
    actividadesVisibles: actividadesVisibles.map(markCompletion),
    listas: getListsDictionary_(),
    areas: AREAS
  };
}

function registrarActividad(payload) {
  validatePayload_(payload);

  const userEmail = Session.getActiveUser().getEmail();
  const coordinador = getCoordinatorByEmail_(userEmail);
  if (!coordinador) {
    throw new Error('No autorizado para registrar actividades.');
  }

  const actividad = getActivityById_(payload.actividadId);
  if (!actividad) {
    throw new Error('La actividad seleccionada no existe.');
  }
  if (actividad.coordinacion !== coordinador.coordinacion) {
    throw new Error('La actividad no pertenece a su coordinación.');
  }

  const isFinalizada = getRecordsByCoordination_(coordinador.coordinacion).some(
    (record) =>
      record.actividadId === payload.actividadId && record.estado === 'Finalizada'
  );
  if (isFinalizada) {
    throw new Error('La actividad ya está finalizada y no admite más registros.');
  }

  const fecha = new Date(payload.fechaActividad);
  const mesSemana = getMonthAndWeek_(fecha);
  const evidencias = uploadEvidenceFiles_(
    payload.fotos || [],
    actividad.indicadorPoa,
    coordinador.coordinacion,
    payload.actividadId
  );

  const registro = {
    timestampRegistro: new Date(),
    registroId: Utilities.getUuid(),
    actividadId: payload.actividadId,
    coordinacion: coordinador.coordinacion,
    correo: userEmail,
    estado: payload.estado,
    fechaActividad: payload.fechaActividad,
    horaInicio: payload.horaInicio,
    horaFin: payload.horaFin,
    mes: mesSemana.mes,
    semanaMes: mesSemana.semana,
    alumnosHombres: Number(payload.alumnosHombres || 0),
    alumnasMujeres: Number(payload.alumnasMujeres || 0),
    docentesHombres: Number(payload.docentesHombres || 0),
    docentesMujeres: Number(payload.docentesMujeres || 0),
    administrativosHombres: Number(payload.administrativosHombres || 0),
    administrativasMujeres: Number(payload.administrativasMujeres || 0),
    tipoActividad: payload.tipoActividad,
    objetivoActividad: payload.objetivoActividad,
    carrerasInvolucradas: payload.carrerasInvolucradas,
    areaPrincipal: payload.areaPrincipal,
    areasApoyo: payload.areasApoyo,
    tipoProtagonista: payload.tipoProtagonista,
    actividadNombre: actividad.actividad,
    indicadorPoa: actividad.indicadorPoa,
    urlsEvidencias: evidencias.join(' | ')
  };

  const docUrl = generarDocumentoActividad_(registro, actividad);
  registro.documentoUrl = docUrl;
  appendObject_(SHEETS.REGISTROS, HEADERS.REGISTROS, registro);
  return { ok: true, registroId: registro.registroId, documento: docUrl, evidencias };
}

function generarDocumentoActividad_(registro, actividad) {
  const fechaDocumento = new Date(registro.fechaActividad);
  const nombreDoc = `ACT_${registro.actividadId}_${registro.fechaActividad}`;
  const carpetaDestino = getReportFolder_(fechaDocumento, registro.coordinacion);
  const doc = DocumentApp.create(nombreDoc);
  const body = doc.getBody();

  body.clear();
  buildDocumentHeader_(body);
  buildGeneralInfoSection_(body, registro);
  buildParticipationSection_(body, registro, actividad);
  buildSimpleTextSection_(body, '3. INDICADOR', registro.indicadorPoa);
  buildSimpleTextSection_(body, '4. OBJETIVO', registro.objetivoActividad);
  buildParticipantsSection_(body, registro);
  buildEvidenceSection_(body, registro.urlsEvidencias);
  buildFooterSection_(body);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(carpetaDestino);

  return doc.getUrl();
}

function buildDocumentHeader_(body) {
  const titulo = body.appendParagraph('Universidad de Ciencias Comerciales');
  titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titulo.setBold(true);
  titulo.setFontSize(18);

  const subtitulo = body.appendParagraph('Reporte de Actividad POA');
  subtitulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subtitulo.setBold(true);
  subtitulo.setFontSize(18);

  const fechaGeneracion = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
  const fechaTexto = body.appendParagraph(`Generado automáticamente: ${fechaGeneracion}`);
  fechaTexto.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fechaTexto.setFontSize(11);
  fechaTexto.setSpacingAfter(14);
}

function buildGeneralInfoSection_(body, registro) {
  appendSectionTitle_(body, '1. INFORMACIÓN GENERAL');
  appendLabelValue_(body, 'Actividad', registro.actividadNombre);
  appendLabelValue_(body, 'Código actividad', registro.actividadId);
  appendLabelValue_(body, 'Realizada por', registro.coordinacion);
  appendLabelValue_(body, 'Fecha', registro.fechaActividad);
  appendLabelValue_(body, 'Hora', `${registro.horaInicio} a ${registro.horaFin}`);
  appendLabelValue_(body, 'Estado', registro.estado);
  body.appendParagraph('');
}

function buildParticipationSection_(body, registro, actividad) {
  appendSectionTitle_(body, '2. PARTICIPACIÓN');
  const coordinaciones = [actividad.areasInvolucradas, actividad.otrasAreas]
    .filter((v) => String(v || '').trim() !== '')
    .join(', ');
  appendLabelValue_(body, 'Coordinaciones participantes', coordinaciones || 'No especificadas');
  appendLabelValue_(body, 'Carreras participantes', registro.carrerasInvolucradas || 'No especificadas');
  appendLabelValue_(body, 'Área principal', registro.areaPrincipal || 'No especificada');
  appendLabelValue_(body, 'Áreas de apoyo', registro.areasApoyo || 'No especificadas');
  body.appendParagraph('');
}

function buildSimpleTextSection_(body, titulo, contenido) {
  appendSectionTitle_(body, titulo);
  const p = body.appendParagraph(String(contenido || 'No especificado'));
  p.setFontSize(11);
  p.setSpacingAfter(10);
}

function buildParticipantsSection_(body, registro) {
  appendSectionTitle_(body, '5. PARTICIPANTES');
  const estudiantesM = Number(registro.alumnasMujeres || 0);
  const estudiantesV = Number(registro.alumnosHombres || 0);
  const docentesM = Number(registro.docentesMujeres || 0);
  const docentesV = Number(registro.docentesHombres || 0);
  const administrativosM = Number(registro.administrativasMujeres || 0);
  const administrativosV = Number(registro.administrativosHombres || 0);
  const totalM = estudiantesM + docentesM + administrativosM;
  const totalV = estudiantesV + docentesV + administrativosV;

  const table = body.appendTable([
    ['Categoría', 'Mujeres', 'Varones'],
    ['Estudiantes', String(estudiantesM), String(estudiantesV)],
    ['Docentes', String(docentesM), String(docentesV)],
    ['Administrativos', String(administrativosM), String(administrativosV)],
    ['TOTAL', String(totalM), String(totalV)]
  ]);
  table.setBorderWidth(1);

  const header = table.getRow(0);
  for (let i = 0; i < header.getNumCells(); i += 1) {
    header.getCell(i).editAsText().setBold(true);
  }
  const totalRow = table.getRow(4);
  for (let j = 0; j < totalRow.getNumCells(); j += 1) {
    totalRow.getCell(j).editAsText().setBold(true);
  }
  body.appendParagraph('');
}

function buildEvidenceSection_(body, urlsEvidencias) {
  appendSectionTitle_(body, '6. EVIDENCIAS');
  const evidencias = splitEvidenceUrls_(urlsEvidencias);
  if (!evidencias.length) {
    const p = body.appendParagraph('No se registraron evidencias.');
    p.setFontSize(11);
    return;
  }

  body.appendParagraph('Evidencias adjuntas').setBold(true).setFontSize(11);
  evidencias.forEach((url) => {
    const fileId = extractDriveFileId_(url);
    if (fileId) {
      try {
        const blob = DriveApp.getFileById(fileId).getBlob();
        body.appendImage(blob).setWidth(420);
        return;
      } catch (error) {}
    }
    body.appendParagraph(url).setLinkUrl(url).setFontSize(10);
  });
  body.appendParagraph('');
}

function buildFooterSection_(body) {
  appendSectionTitle_(body, '7. PIE');
  body.appendParagraph('Documento generado automáticamente por Sistema POA.').setFontSize(10);
}

function appendSectionTitle_(body, title) {
  const p = body.appendParagraph(title);
  p.setBold(true);
  p.setFontSize(14);
  p.setForegroundColor('#0B5394');
  p.setSpacingBefore(8);
  p.setSpacingAfter(4);
}

function appendLabelValue_(body, label, value) {
  const p = body.appendParagraph(`${label}: ${String(value || 'No especificado')}`);
  p.setFontSize(11);
  p.setSpacingAfter(2);
}

function splitEvidenceUrls_(urlsEvidencias) {
  return String(urlsEvidencias || '')
    .split('|')
    .map((url) => String(url || '').trim())
    .filter((url) => url !== '');
}

function extractDriveFileId_(url) {
  const text = String(url || '').trim();
  if (!text) {
    return '';
  }
  const directIdMatch = text.match(/[-\w]{25,}/);
  if (text.indexOf('/d/') !== -1) {
    const pathMatch = text.match(/\/d\/([-\w]{25,})/);
    return pathMatch ? pathMatch[1] : '';
  }
  const queryMatch = text.match(/[?&]id=([-\w]{25,})/);
  if (queryMatch) {
    return queryMatch[1];
  }
  return directIdMatch ? directIdMatch[0] : '';
}

function getReportFolder_(fechaActividad, coordinacion) {
  const reportesRoot = getOrCreateReportsRootFolder_();
  const anio = Utilities.formatDate(fechaActividad, Session.getScriptTimeZone(), 'yyyy');
  const mes = Utilities.formatDate(fechaActividad, Session.getScriptTimeZone(), 'MMMM');
  const anioFolder = getOrCreateSubfolder_(reportesRoot, anio);
  const coordinacionFolder = getOrCreateSubfolder_(anioFolder, sanitizeFolderName_(coordinacion));
  return getOrCreateSubfolder_(coordinacionFolder, mes);
}

function getOrCreateReportsRootFolder_() {
  const baseFolder = DRIVE_ROOT_FOLDER_ID !== 'REEMPLAZAR_CON_FOLDER_ID_PRINCIPAL'
    ? DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID)
    : DriveApp.getRootFolder();
  return getOrCreateSubfolder_(baseFolder, 'POA_REPORTES');
}

function sanitizeFolderName_(value) {
  return String(value || 'Sin coordinacion')
    .replace(/[\\/:*?"<>|#]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getCoordinatorByEmail_(email) {
  const rows = getSheetObjects_(SHEETS.COORDINADORES, HEADERS.COORDINADORES);
  return (
    rows.find(
      (r) =>
        String(r.correo).trim().toLowerCase() ===
          String(email).trim().toLowerCase() &&
        String(r.activo).toLowerCase() !== 'false'
    ) || null
  );
}

function getActivitiesByCoordination_(coordinacion) {
  return getSheetObjects_(SHEETS.ACTIVIDADES, HEADERS.ACTIVIDADES).filter(
    (row) => row.coordinacion === coordinacion
  );
}

function getVisibleActivities_(coordinacion, cuatrimestre) {
  const normalizedCoord = normalizeText_(coordinacion);
  const normalizedQuarter = normalizeQuarter_(cuatrimestre, getCurrentQuarter_());
  return getSheetObjects_(SHEETS.ACTIVIDADES, HEADERS.ACTIVIDADES)
    .map((row) => {
      const owner = String(row.coordinacion || '').trim();
      const involvedAreas = splitAndNormalizeList_(row.areasInvolucradas);
      const isOwner = normalizeText_(owner) === normalizedCoord;
      const isInvolved = involvedAreas.includes(normalizedCoord);
      return {
        ...row,
        cuatrimestre: String(row.cuatrimestre || '').trim(),
        coordinacion: owner,
        areasInvolucradasLista: splitList_(row.areasInvolucradas),
        otrasAreasLista: splitList_(row.otrasAreas),
        esPropietario: isOwner,
        esParticipante: !isOwner && isInvolved
      };
    })
    .filter((row) => (row.esPropietario || row.esParticipante) && String(row.cuatrimestre || '').trim() === String(normalizedQuarter));
}

function getActivityById_(actividadId) {
  return (
    getSheetObjects_(SHEETS.ACTIVIDADES, HEADERS.ACTIVIDADES).find(
      (row) => row.actividadId === actividadId
    ) || null
  );
}

function getRecordsByCoordination_(coordinacion) {
  return getSheetObjects_(SHEETS.REGISTROS, HEADERS.REGISTROS).filter(
    (row) => row.coordinacion === coordinacion
  );
}

function getListsDictionary_() {
  const requiredLists = ['tipoActividad', 'tipoProtagonista', 'indicadorPoa', 'indicadorEstrategia', 'carreras'];
  const defaultError = 'Error list.';
  const result = requiredLists.reduce((acc, key) => {
    acc[key] = { items: [], error: null };
    return acc;
  }, {});

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.LISTAS);
  if (!sh) {
    requiredLists.forEach((key) => {
      result[key].error = defaultError;
    });
    return result;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    requiredLists.forEach((key) => {
      result[key].error = defaultError;
    });
    return result;
  }

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map((v) => String(v || '').trim());
  const dataRows = values.slice(1);

  requiredLists.forEach((key) => {
    const colIndex = headers.indexOf(key);
    if (colIndex < 0) {
      result[key].error = defaultError;
      return;
    }

    const items = dataRows
      .map((row) => String(row[colIndex] || '').trim())
      .filter((value) => value !== '');

    result[key].items = items;
  });

  result.indicadoresDetalle = getIndicatorDetails_(values);

  return result;
}

function getIndicatorDetails_(sheetValues) {
  if (!Array.isArray(sheetValues) || !sheetValues.length) {
    return [];
  }

  const headers = sheetValues[0].map((v) => String(v || '').trim());
  const indicatorCol = headers.indexOf('indicadorPoa');
  if (indicatorCol < 0) {
    return [];
  }
  const strategyCol = headers.indexOf('indicadorEstrategia');

  return sheetValues.slice(1).reduce((acc, row) => {
    const indicador = String(row[indicatorCol] || '').trim();
    const numberMatch = indicador.match(/^\d+/);
    if (!indicador || !numberMatch) {
      return acc;
    }
    acc.push({
      numero: numberMatch[0],
      indicador,
      indicadorEstrategia:
        strategyCol >= 0 ? String(row[strategyCol] || '').trim() : ''
    });
    return acc;
  }, []);
}

function actualizarOtrasAreasActividad(payload) {
  const actividadId = String((payload && payload.actividadId) || '').trim();
  const nuevasAreas = Array.isArray(payload && payload.otrasAreas)
    ? payload.otrasAreas.map((a) => String(a || '').trim()).filter((a) => a !== '')
    : [];
  if (!actividadId) {
    throw new Error('Debe seleccionar una actividad.');
  }

  const userEmail = Session.getActiveUser().getEmail();
  const coordinador = getCoordinatorByEmail_(userEmail);
  if (!coordinador) {
    throw new Error('No autorizado.');
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ACTIVIDADES);
  if (!sh) {
    throw new Error(`No existe la hoja ${SHEETS.ACTIVIDADES}.`);
  }

  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
  const idxActividad = headerRow.indexOf('actividadId');
  const idxCoordinacion = headerRow.indexOf('coordinacion');
  const idxAreas = headerRow.indexOf('areasInvolucradas');
  const idxOtrasAreas = headerRow.indexOf('otrasAreas');
  if (idxActividad < 0 || idxCoordinacion < 0 || idxAreas < 0 || idxOtrasAreas < 0) {
    throw new Error('La hoja ActividadesPOA debe tener las columnas actividadId, coordinacion, areasInvolucradas y otrasAreas.');
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) {
    throw new Error('No hay actividades para actualizar.');
  }
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const normalizedCoord = normalizeText_(coordinador.coordinacion);
  const rowIndex = data.findIndex((row) => String(row[idxActividad] || '').trim() === actividadId);
  if (rowIndex < 0) {
    throw new Error('No se encontró la actividad seleccionada.');
  }

  const ownerCoord = String(data[rowIndex][idxCoordinacion] || '').trim();
  if (normalizeText_(ownerCoord) !== normalizedCoord) {
    throw new Error('Solo el dueño de la actividad puede agregar áreas de apoyo.');
  }

  const existentes = splitList_(data[rowIndex][idxAreas]);
  const existentesNormalized = new Set(existentes.map((a) => normalizeText_(a)));
  const limpias = nuevasAreas.filter((area) => !existentesNormalized.has(normalizeText_(area)));
  const valorFinal = Array.from(new Set(limpias)).join(', ');

  sh.getRange(rowIndex + 2, idxOtrasAreas + 1).setValue(valorFinal);
  return { ok: true, actividadId, otrasAreas: splitList_(valorFinal) };
}

function splitAndNormalizeList_(value) {
  return splitList_(value).map((v) => normalizeText_(v));
}

function splitList_(value) {
  return String(value || '')
    .split(',')
    .map((v) => String(v || '').trim())
    .filter((v) => v !== '');
}

function normalizeText_(value) {
  return String(value || '').trim().toLowerCase();
}

function getCurrentQuarter_() {
  const month = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M'));
  return Math.ceil(month / 4);
}

function normalizeQuarter_(value, fallbackQuarter) {
  const raw = Number(String(value || '').trim());
  if ([1, 2, 3, 4].indexOf(raw) >= 0) {
    return raw;
  }
  return fallbackQuarter;
}

function uploadEvidenceFiles_(files, indicadorPoa, coordinacion, actividadId) {
  if (!files.length) {
    return [];
  }
  if (DRIVE_ROOT_FOLDER_ID === 'REEMPLAZAR_CON_FOLDER_ID_PRINCIPAL') {
    throw new Error('Configurar DRIVE_ROOT_FOLDER_ID antes de subir evidencias.');
  }

  const root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  const indicatorFolder = getOrCreateSubfolder_(root, String(indicadorPoa));
  const coordinationFolder = getOrCreateSubfolder_(indicatorFolder, coordinacion);

  return files.map((file) => {
    const content = Utilities.base64Decode(file.base64);
    const blob = Utilities.newBlob(content, file.mimeType, file.fileName);
    const saved = coordinationFolder.createFile(blob);
    saved.setDescription(`Actividad ${actividadId}`);
    return saved.getUrl();
  });
}

function getOrCreateSubfolder_(parentFolder, folderName) {
  const existing = parentFolder.getFoldersByName(folderName);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(folderName);
}

function getMonthAndWeek_(dateValue) {
  const monthName = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MMMM');
  const day = Number(Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'd'));
  const semana = `Semana ${Math.ceil(day / 7)}`;
  return {
    mes: monthName.charAt(0).toUpperCase() + monthName.slice(1),
    semana
  };
}

function validatePayload_(payload) {
  const required = [
    'actividadId',
    'estado',
    'fechaActividad',
    'horaInicio',
    'horaFin',
    'tipoActividad',
    'objetivoActividad',
    'areaPrincipal'
  ];

  required.forEach((field) => {
    if (!payload[field]) {
      throw new Error(`El campo ${field} es obligatorio.`);
    }
  });
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const sameHeaders = headers.every((h, i) => firstRow[i] === h);
  if (!sameHeaders) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getSheetObjects_(sheetName, headers) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) {
    throw new Error(`No existe la hoja ${sheetName}. Ejecute initializeSheets().`);
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function appendObject_(sheetName, headers, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) {
    throw new Error(`No existe la hoja ${sheetName}.`);
  }
  const row = headers.map((header) => obj[header] || '');
  sh.appendRow(row);
}
