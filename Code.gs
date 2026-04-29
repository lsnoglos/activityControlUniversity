const SHEETS = {
  ACTIVIDADES: 'ActividadesPOA',
  COORDINADORES: 'Coordinadores',
  REGISTROS: 'Registros',
  LISTAS: 'Listas'
};

const DRIVE_ROOT_FOLDER_ID = 'REEMPLAZAR_CON_FOLDER_ID_PRINCIPAL';
const DRIVE_ROOT_FOLDER_NAME_PREFIX = 'EVIDENCIAS POA';
const INSTITUTIONAL_BLUE = '#1F4E78';
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
  COORDINADORES: ['coordinacion', 'correo', 'nombre', 'activo'],
  REGISTROS: [
    'timestampRegistro',
    'registroId',
    'actividadId',
    'coordinacion',
    'correo',
    'coordinadorNombre',
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
    'areasApoyo',
    'tipoProtagonista',
    'actividadNombre',
    'indicadorPoa',
    'urlsEvidencias',
    'documentoUrl',
    'pdfUrl'
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
    coordinadorNombre: coordinador.nombre || '',
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
  const areaNames = getParticipantAreaNames_(actividad, coordinador.coordinacion);
  const participantEmails = getParticipantEmails_(areaNames);
  const storage = prepareActivityStorage_(registroFolderNames_(actividad, payload), participantEmails);
  const evidencias = uploadEvidenceFiles_(payload.fotos || [], storage.evidenceFolder, payload.actividadId);

  const registro = {
    timestampRegistro: new Date(),
    registroId: Utilities.getUuid(),
    actividadId: payload.actividadId,
    coordinacion: coordinador.coordinacion,
    correo: userEmail,
    coordinadorNombre: coordinador.nombre || '',
    estado: resolveStatusForRecord_(coordinador.coordinacion, payload.actividadId),
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
    areasApoyo: payload.areasApoyo,
    tipoProtagonista: payload.tipoProtagonista,
    actividadNombre: actividad.actividad,
    indicadorPoa: actividad.indicadorPoa,
    urlsEvidencias: evidencias.join(' | '),
    documentoUrl: '',
    pdfUrl: ''
  };

  const documentos = generarDocumentoActividad_(registro, actividad, storage, participantEmails);
  registro.documentoUrl = documentos.documentoUrl || '';
  registro.pdfUrl = documentos.pdfUrl || '';

  appendObject_(SHEETS.REGISTROS, HEADERS.REGISTROS, registro);
  return {
    ok: true,
    registroId: registro.registroId,
    evidencias,
    documento: registro.documentoUrl,
    pdf: registro.pdfUrl
  };
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


function resolveStatusForRecord_(coordinacion, actividadId) {
  const hasPending = getRecordsByCoordination_(coordinacion).some(
    (record) => record.actividadId === actividadId && record.estado === 'Pendiente'
  );
  return hasPending ? 'Finalizada' : 'Pendiente';
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

function registroFolderNames_(actividad, payload) {
  const dateValue = new Date(payload.fechaActividad);
  const year = Utilities.formatDate(
    String(dateValue) === 'Invalid Date' ? new Date() : dateValue,
    Session.getScriptTimeZone(),
    'yyyy'
  );
  return {
    evidenciasPoaFolder: `${DRIVE_ROOT_FOLDER_NAME_PREFIX} ${year}`,
    indicadorLabel: `Indicador ${extractIndicatorNumber_(actividad.indicadorPoa)}`,
    actividadNombre: sanitizeFolderName_(actividad.actividad || payload.actividadId || 'Actividad')
  };
}

function prepareActivityStorage_(names, participantEmails) {
  const root = getConfiguredRootFolder_();
  const evidenciasRoot = getOrCreateSubfolder_(root, names.evidenciasPoaFolder);
  const indicatorFolder = getOrCreateSubfolder_(evidenciasRoot, names.indicadorLabel);
  const activityFolder = getOrCreateSubfolder_(indicatorFolder, names.actividadNombre);
  const evidenceFolder = getOrCreateSubfolder_(activityFolder, 'Evidencias');

  applyEditorsToFolder_(activityFolder, participantEmails);
  applyEditorsToFolder_(evidenceFolder, participantEmails);

  return {
    activityFolder,
    evidenceFolder,
    reportFolder: activityFolder
  };
}

function uploadEvidenceFiles_(files, evidenceFolder, actividadId) {
  if (!files.length) {
    return [];
  }
  return files.map((file) => {
    const content = Utilities.base64Decode(file.base64);
    const blob = Utilities.newBlob(content, file.mimeType, file.fileName);
    const saved = evidenceFolder.createFile(blob);
    saved.setDescription(`Actividad ${actividadId}`);
    return saved.getUrl();
  });
}

function generarDocumentoActividad_(registro, actividad, storage, participantEmails) {
  const dateIso = formatDateForFileName_(registro.fechaActividad);
  const docName = `ACT_${registro.actividadId}_${dateIso}`;
  const docFile = DocumentApp.create(docName);
  const doc = DocumentApp.openById(docFile.getId());
  const body = doc.getBody();
  buildActivityDocument_(body, registro, actividad);
  doc.saveAndClose();

  const source = DriveApp.getFileById(docFile.getId());
  const copyInReport = source.makeCopy(docName, storage.reportFolder);
  const copyInActivity = source.makeCopy(docName, storage.activityFolder);
  source.setTrashed(true);

  applyEditorsToFile_(copyInReport, participantEmails);
  applyEditorsToFile_(copyInActivity, participantEmails);

  const pdfBlob = DocumentApp.openById(copyInActivity.getId()).getAs(MimeType.PDF).setName(`${docName}.pdf`);
  const pdfInActivity = storage.activityFolder.createFile(pdfBlob);
  const pdfInReport = storage.reportFolder.createFile(pdfBlob.copyBlob().setName(`${docName}.pdf`));
  applyEditorsToFile_(pdfInActivity, participantEmails);
  applyEditorsToFile_(pdfInReport, participantEmails);

  return {
    documentoUrl: copyInActivity.getUrl(),
    pdfUrl: pdfInActivity.getUrl(),
    documentoReporteUrl: copyInReport.getUrl()
  };
}

function getOrCreateSubfolder_(parentFolder, folderName) {
  const existing = parentFolder.getFoldersByName(folderName);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(folderName);
}

function buildActivityDocument_(body, registro, actividad) {
  body.clear();
  body.setAttributes({
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
    [DocumentApp.Attribute.FONT_SIZE]: 11
  });

  appendTitle_(body, 'Universidad de Ciencias Comerciales');
  appendTitle_(body, 'Reporte de Actividad POA');
  appendNormalLine_(body, `Fecha de generación: ${formatDateTimeNow_()}`, true);
  body.appendParagraph('');

  appendSectionTitle_(body, '1. INFORMACIÓN GENERAL');
  appendLabelValue_(body, 'Actividad', registro.actividadNombre);
  appendLabelValue_(body, 'Código actividad', registro.actividadId);
  appendLabelValue_(body, 'Realizada por', registro.coordinacion);
  appendLabelValue_(body, 'Coordinador responsable', registro.coordinadorNombre || 'N/D');
  appendLabelValue_(body, 'Fecha', registro.fechaActividad);
  appendLabelValue_(body, 'Hora', `${registro.horaInicio} a ${registro.horaFin}`);
  appendLabelValue_(body, 'Estado', registro.estado);

  appendSectionTitle_(body, '2. PARTICIPACIÓN');
  appendLabelValue_(body, 'Coordinaciones participantes', joinNonEmpty_([actividad.areasInvolucradas, actividad.otrasAreas], ' | '));
  appendLabelValue_(body, 'Carreras participantes', registro.carrerasInvolucradas);
  appendLabelValue_(body, 'Áreas de apoyo', registro.areasApoyo);

  appendSectionTitle_(body, '3. INDICADOR');
  appendNormalLine_(body, registro.indicadorPoa || 'N/D');

  appendSectionTitle_(body, '4. OBJETIVO');
  appendNormalLine_(body, registro.objetivoActividad || 'N/D');

  appendSectionTitle_(body, '5. PARTICIPANTES');
  appendParticipantsTable_(body, registro);

  appendSectionTitle_(body, '6. EVIDENCIAS');
  appendEvidenceSection_(body, splitPipeList_(registro.urlsEvidencias));

  appendSectionTitle_(body, '7. PIE');
  appendNormalLine_(body, 'Documento generado automáticamente por Sistema POA.');
}

function appendTitle_(body, text) {
  const p = body.appendParagraph(text || '');
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.setBold(true);
  p.setFontSize(18);
}

function appendSectionTitle_(body, text) {
  const p = body.appendParagraph(text || '');
  p.setBold(true);
  p.setFontSize(14);
  p.setForegroundColor(INSTITUTIONAL_BLUE);
  p.setSpacingBefore(14);
  p.setSpacingAfter(6);
}

function appendLabelValue_(body, label, value) {
  const p = body.appendParagraph(`${label}: ${value || 'N/D'}`);
  p.setFontSize(11);
  p.setSpacingAfter(4);
}

function appendNormalLine_(body, text, isMuted) {
  const p = body.appendParagraph(text || '');
  p.setFontSize(11);
  if (isMuted) {
    p.setForegroundColor('#5f6368');
  }
}

function appendParticipantsTable_(body, registro) {
  const alumnosM = Number(registro.alumnosHombres || 0);
  const alumnasF = Number(registro.alumnasMujeres || 0);
  const docentesM = Number(registro.docentesHombres || 0);
  const docentesF = Number(registro.docentesMujeres || 0);
  const admM = Number(registro.administrativosHombres || 0);
  const admF = Number(registro.administrativasMujeres || 0);
  const totalM = alumnosM + docentesM + admM;
  const totalF = alumnasF + docentesF + admF;

  const table = body.appendTable([
    ['Categoría', 'Mujeres', 'Varones'],
    ['Estudiantes', String(alumnasF), String(alumnosM)],
    ['Docentes', String(docentesF), String(docentesM)],
    ['Administrativos', String(admF), String(admM)],
    ['TOTAL', String(totalF), String(totalM)]
  ]);

  table.setBorderWidth(1);
  const headerRow = table.getRow(0);
  for (let c = 0; c < headerRow.getNumCells(); c++) {
    headerRow.getCell(c).setBackgroundColor('#E8EEF7').editAsText().setBold(true);
  }
}

function appendEvidenceSection_(body, evidenceUrls) {
  if (!evidenceUrls.length) {
    appendNormalLine_(body, 'No se registraron evidencias.');
    return;
  }
  appendNormalLine_(body, 'Evidencias adjuntas');
  evidenceUrls.forEach((url, index) => {
    const fileId = extractDriveFileId_(url);
    if (fileId) {
      try {
        const imageBlob = DriveApp.getFileById(fileId).getBlob();
        body.appendImage(imageBlob).setWidth(450);
        return;
      } catch (error) {
        // Si falla inserción de imagen, dejar enlace.
      }
    }
    const p = body.appendParagraph(`Evidencia ${index + 1}`);
    p.setLinkUrl(url);
    p.setForegroundColor('#1155CC');
  });
}

function getConfiguredRootFolder_() {
  if (DRIVE_ROOT_FOLDER_ID === 'REEMPLAZAR_CON_FOLDER_ID_PRINCIPAL') {
    throw new Error('Configurar DRIVE_ROOT_FOLDER_ID antes de registrar actividades con evidencias/documentos.');
  }
  return DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
}

function applyEditorsToFolder_(folder, emails) {
  const uniqueEmails = uniqueEmails_(emails);
  uniqueEmails.forEach((email) => {
    try {
      folder.addEditor(email);
    } catch (error) {
      Logger.log(`No fue posible agregar editor ${email} a carpeta ${folder.getName()}: ${error}`);
    }
  });
}

function applyEditorsToFile_(file, emails) {
  const uniqueEmails = uniqueEmails_(emails);
  uniqueEmails.forEach((email) => {
    try {
      file.addEditor(email);
    } catch (error) {
      Logger.log(`No fue posible agregar editor ${email} a archivo ${file.getName()}: ${error}`);
    }
  });
}

function getParticipantAreaNames_(actividad, ownerCoordination) {
  const areas = []
    .concat(splitList_(actividad.areasInvolucradas))
    .concat(splitList_(actividad.otrasAreas))
    .concat([ownerCoordination]);
  return Array.from(new Set(areas.map((v) => String(v || '').trim()).filter((v) => v !== '')));
}

function getParticipantEmails_(areas) {
  const coordinadores = getSheetObjects_(SHEETS.COORDINADORES, HEADERS.COORDINADORES);
  const normalizedAreas = new Set(areas.map((area) => normalizeText_(area)));
  const emails = coordinadores
    .filter((c) => String(c.activo).toLowerCase() !== 'false' && normalizedAreas.has(normalizeText_(c.coordinacion)))
    .map((c) => String(c.correo || '').trim().toLowerCase())
    .filter((email) => email !== '');
  return Array.from(new Set(emails));
}

function uniqueEmails_(emails) {
  return Array.from(
    new Set((emails || []).map((e) => String(e || '').trim().toLowerCase()).filter((e) => e !== ''))
  );
}

function extractIndicatorNumber_(indicadorPoa) {
  const raw = String(indicadorPoa || '').trim();
  const numberMatch = raw.match(/\d+/);
  return numberMatch ? numberMatch[0] : 'Sin indicador';
}

function sanitizeFolderName_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%{}~]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function formatDateForFileName_(value) {
  const dateValue = new Date(value);
  if (String(dateValue) === 'Invalid Date') {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function splitPipeList_(value) {
  return String(value || '')
    .split('|')
    .map((v) => String(v || '').trim())
    .filter((v) => v !== '');
}

function joinNonEmpty_(values, separator) {
  return (values || [])
    .map((v) => String(v || '').trim())
    .filter((v) => v !== '')
    .join(separator || ', ');
}

function extractDriveFileId_(url) {
  const raw = String(url || '').trim();
  if (!raw) {
    return null;
  }
  const idByD = raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (idByD && idByD[1]) {
    return idByD[1];
  }
  const idByParam = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idByParam && idByParam[1]) {
    return idByParam[1];
  }
  return null;
}

function formatDateTimeNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
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
    'fechaActividad',
    'horaInicio',
    'horaFin',
    'tipoActividad',
    'objetivoActividad'
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
  const row = headers.map((header) => {
    const value = obj[header];
    return value === undefined || value === null ? '' : value;
  });
  sh.appendRow(row);
}
