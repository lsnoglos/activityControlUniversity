# Sistema de Control de Actividades POA (Google Apps Script)

Este repositorio contiene una base funcional para gestionar actividades POA por coordinación en Google Sheets + Apps Script.

## Qué hace

- Autentica al coordinador por correo institucional (`Session.getActiveUser().getEmail()`).
- Muestra actividades **pendientes** y **realizadas** por coordinación.
- Permite registrar actividades con estado automático (`Pendiente`/`Finalizada`) y:
  - fecha, hora inicio, hora fin,
  - mes y semana (calculados automáticamente),
  - participantes por sexo y rol,
  - tipo de actividad, objetivo, carreras, áreas de apoyo, tipo de protagonista (multiselección),
  - evidencias fotográficas.
- Guarda registros en la hoja `Registros`.
- Organiza evidencias en Google Drive:
  - Carpeta raíz configurable.
  - Carpeta anual `EVIDENCIAS POA {AÑO}`.
  - Subcarpeta por indicador con formato `Indicador {número}`.
  - Subcarpeta por nombre de actividad, con `Evidencias` y documentos.

## Estructura de hojas

La función `initializeSheets()` crea (si no existen) estas hojas y cabeceras:

1. `ActividadesPOA`
   - `anio`, `actividadId`, `coordinacion`, `actividad`, `indicadorPoa`, `ejeEne`, `areasInvolucradas`, `cuatrimestre`
2. `Coordinadores`
   - `coordinacion`, `correo`, `nombre`, `activo`
3. `Registros`
   - Campos de control, tiempos, participantes, metadatos de actividad y URLs de evidencias.
4. `Listas`
   - Formato por columnas (recomendado): una columna por lista.
   - Cabeceras iniciales sugeridas por `initializeSheets()`:
     - `tipoActividad`, `tipoProtagonista`, `indicadorPoa`, `codigoEstrategia`
   - También se mantiene compatibilidad con el formato anterior `lista`, `valor`.

## Áreas institucionales

El sistema incluye un catálogo fijo de áreas para el registro de actividades:

- Ingeniería Civil
- CCEEyJJ
- Investigación
- Posgrado y EC
- BE - Proy. Social
- Biblioteca
- Supervisión Metodológica
- Dirección Académica
- Comunicación Institucional
- Recursos Humanos
- Registro Académico
- Ingeniería Agronómica
- Diseño Gráfico y Arq
- Gestión de Calidad
- TIC

Regla de uso: cada actividad puede incluir **múltiples áreas de apoyo** gestionadas por el dueño de la actividad.

## Listas esperadas en `Listas`

Use una columna por cada lista y cargue sus valores hacia abajo en la misma columna. Ejemplo:

- `tipoActividad`: `Curricular`, `Formación continua`, `Participación en programas nacionales`, `Tecnologías`.
- `tipoProtagonista`: `Estudiante`, `Docente`, `Personal administrativo`, `Servidores públicos`, etc.
- `indicadorPoa`: descripciones de indicadores.
- `codigoEstrategia`: códigos de estrategia asociados.

## Configuración rápida

1. Crear un proyecto de Apps Script vinculado a una hoja de cálculo.
2. Copiar `Code.gs`, `Index.html` y `appsscript.json`.
3. En `Code.gs`, configurar:
   - `DRIVE_ROOT_FOLDER_ID` con el ID de carpeta raíz para evidencias.
4. Ejecutar manualmente `initializeSheets()` una vez.
5. Cargar datos en `ActividadesPOA`, `Coordinadores` y `Listas`.
6. Desplegar como **Web App**:
   - Ejecutar como: *Usuario que accede*.
   - Acceso: según política institucional (ideal: dominio institucional).

## Notas importantes

- Para que `Session.getActiveUser().getEmail()` retorne correo, la app debe operar bajo políticas de dominio (Workspace).
- Si el correo no existe en `Coordinadores` o está inactivo (`activo = false`), el acceso se bloquea.
- Si no configura `DRIVE_ROOT_FOLDER_ID`, la carga de fotos falla por diseño.
