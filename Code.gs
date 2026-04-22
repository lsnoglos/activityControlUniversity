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
    'urlsEvidencias'
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
  ensureSheetWithHeaders_(ss, SHEETS.LISTAS, ['tipoActividad', 'tipoProtagonista', 'indicadorPoa', 'codigoEstrategia']);
}

function getInitialData() {
  const userEmail = Session.getActiveUser().getEmail();
  const coordinador = getCoordinatorByEmail_(userEmail);
  if (!coordinador) {
    return {
      authorized: false,
      userEmail,
      message: `El correo ${userEmail} no tiene acceso.`
    };
  }

  const actividades = getActivitiesByCoordination_(coordinador.coordinacion);
  const registros = getRecordsByCoordination_(coordinador.coordinacion);
  const actividadIdsCompletadas = new Set(
    registros.filter((r) => r.estado === 'Finalizada').map((r) => r.actividadId)
  );

  const actividadesCompletadas = actividades.filter((a) =>
    actividadIdsCompletadas.has(a.actividadId)
  );
  const actividadesPendientes = actividades.filter(
    (a) => !actividadIdsCompletadas.has(a.actividadId)
  );

  return {
    authorized: true,
    userEmail,
    coordinacion: coordinador.coordinacion,
    actividadesPendientes,
    actividadesCompletadas,
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

  appendObject_(SHEETS.REGISTROS, HEADERS.REGISTROS, registro);
  return { ok: true, registroId: registro.registroId, evidencias };
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
  const requiredLists = ['tipoActividad', 'tipoProtagonista', 'indicadorPoa', 'codigoEstrategia'];
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

  return result;
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
