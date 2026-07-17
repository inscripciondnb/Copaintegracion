/**
 * COPA INTEGRACIÓN DNB 2026 — Backend final
 * Guarda inscripciones, participantes, hospedaje y dorsales.
 * Genera una ficha Word editable y envía confirmación por correo.
 */

const CARPETA_RAIZ = 'Copa Integración DNB 2026 — Archivos';
const MAIL_ORGANIZACION = ''; // Completar con el correo que recibirá cada ficha Word.
const LOGO_FILE_ID = '';      // Opcional: ID del escudo guardado en Drive.

function doGet() {
  return respuesta({ ok: true, servicio: 'Copa Integración DNB 2026' });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No se recibieron datos.');
    }

    const d = JSON.parse(e.postData.contents);
    validarDatos(d);

    const id = generarId();
    const carpeta = crearCarpetaInscripcion(id, d.capitan.nombre);
    const participantes = guardarArchivosParticipantes(d, carpeta);

    guardarInscripcion(id, d, participantes);
    guardarParticipantes(id, d, participantes);
    guardarHospedaje(id, d, participantes);
    guardarDorsales(id, d, participantes);

    const word = generarWord(id, d, participantes, carpeta);
    enviarConfirmacion(id, d, participantes);
    enviarWordOrganizacion(id, d, word, carpeta);

    return respuesta({ ok: true, id: id });
  } catch (err) {
    console.error(err);
    return respuesta({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function validarDatos(d) {
  if (!d || !d.categoria) throw new Error('Falta la modalidad seleccionada.');
  if (!d.capitan || !d.capitan.nombre || !d.capitan.email) throw new Error('Faltan datos del capitán.');
  if (!d.capitan.hospedaje) throw new Error('Falta seleccionar el hospedaje del capitán.');
  if (!Array.isArray(d.competidores) || !d.competidores.length) throw new Error('No hay competidores cargados.');

  d.competidores.forEach(function(c, i) {
    if (!c.nombre || !c.sexo || !c.documento || !c.hospedaje) {
      throw new Error('Faltan datos obligatorios del competidor ' + (i + 1) + '.');
    }
  });
}

function generarId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const n = parseInt(props.getProperty('contador') || '0', 10) + 1;
    props.setProperty('contador', String(n));
    return 'COPA-' + ('0000' + n).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function crearCarpetaInscripcion(id, nombreCapitan) {
  const it = DriveApp.getFoldersByName(CARPETA_RAIZ);
  const raiz = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_RAIZ);
  return raiz.createFolder(id + ' — ' + limpiarNombre(nombreCapitan));
}

function limpiarNombre(texto) {
  return String(texto || '').replace(/[\\/:*?"<>|]/g, '-').trim();
}

function guardarArchivosParticipantes(d, carpeta) {
  const lista = [];

  lista.push({
    numero: 0,
    dorsal: '',
    rol: 'Capitán',
    nombre: d.capitan.nombre,
    documento: d.capitan.ci,
    celular: d.capitan.celular,
    email: d.capitan.email,
    sexo: d.capitan.sexo || '',
    edad: d.capitan.edad || '',
    hospedaje: d.capitan.hospedaje,
    carnetSalud: null,
    ergometria: null,
    linkCarnet: '',
    linkErgo: ''
  });

  (d.competidores || []).forEach(function(c, i) {
    const fila = Object.assign({}, c);
    fila.numero = i + 1;
    fila.rol = 'Competidor';
    fila.linkCarnet = guardarArchivo(c.carnetSalud, carpeta, 'Competidor_' + (i + 1) + '_' + limpiarNombre(c.nombre) + '_Carnet');
    fila.linkErgo = guardarArchivo(c.ergometria, carpeta, 'Competidor_' + (i + 1) + '_' + limpiarNombre(c.nombre) + '_Ergometria');
    lista.push(fila);
  });

  return lista;
}

function guardarArchivo(archivo, carpeta, nombreBase) {
  if (!archivo || !archivo.base64) return '';
  const ext = (archivo.nombre && archivo.nombre.match(/\.[^.]+$/) || [''])[0];
  const blob = Utilities.newBlob(
    Utilities.base64Decode(archivo.base64),
    archivo.tipo || 'application/octet-stream',
    nombreBase + ext
  );
  return carpeta.createFile(blob).getUrl();
}

function hoja(nombre, encabezados) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(nombre);
  if (!h) h = libro.insertSheet(nombre);

  if (h.getLastRow() === 0) {
    h.appendRow(encabezados);
    h.getRange(1, 1, 1, encabezados.length)
      .setFontWeight('bold')
      .setBackground('#A61E24')
      .setFontColor('#FFFFFF')
      .setWrap(true);
    h.setFrozenRows(1);
  }
  return h;
}

function guardarInscripcion(id, d, participantes) {
  hoja('INSCRIPCIONES', [
    'N.º inscripción', 'Fecha', 'Hora', 'Modalidad seleccionada',
    'Capitán', 'C.I. capitán', 'Teléfono', 'Correo', 'Total personas',
    'Solicitan hospedaje', 'No solicitan hospedaje'
  ]).appendRow([
    id,
    formatearFecha(new Date(), 'dd/MM/yyyy'),
    formatearFecha(new Date(), 'HH:mm'),
    d.categoria,
    d.capitan.nombre,
    d.capitan.ci,
    d.capitan.celular,
    d.capitan.email,
    participantes.length,
    participantes.filter(solicitaHospedaje).length,
    participantes.filter(function(p) { return !solicitaHospedaje(p); }).length
  ]);
}

function guardarParticipantes(id, d, participantes) {
  const h = hoja('PARTICIPANTES', [
    'N.º inscripción', 'N.º', 'Rol', 'Nombre y apellido', 'C.I./Documento',
    'Sexo', 'Edad', 'Modalidad seleccionada', 'Solicita hospedaje',
    'Carnet de salud', 'Ergometría'
  ]);

  participantes.forEach(function(p) {
    h.appendRow([
      id,
      p.numero === 0 ? 'CAP' : p.numero,
      p.rol,
      p.nombre,
      p.documento || '',
      p.sexo || '',
      p.edad || '',
      d.categoria,
      solicitaHospedaje(p) ? 'SÍ' : 'NO',
      p.linkCarnet || (p.rol === 'Capitán' ? '—' : 'FALTA'),
      p.linkErgo || (p.rol === 'Capitán' ? '—' : (necesitaErgo(p) ? 'FALTA' : 'No requiere'))
    ]);
  });
}

function guardarHospedaje(id, d, participantes) {
  const h = hoja('HOSPEDAJE', [
    'N.º inscripción', 'Nombre y apellido', 'Rol', 'Sexo',
    'Modalidad seleccionada', 'Solicita hospedaje', 'Cama requerida'
  ]);

  participantes.forEach(function(p) {
    h.appendRow([
      id,
      p.nombre,
      p.rol,
      p.sexo || '',
      d.categoria,
      solicitaHospedaje(p) ? 'SÍ' : 'NO',
      solicitaHospedaje(p) ? 1 : 0
    ]);
  });
}

function guardarDorsales(id, d, participantes) {
  const h = hoja('DORSALES', [
    'N.º inscripción', 'N.º dorsal', 'Nombre y apellido', 'Sexo',
    'Modalidad seleccionada', 'Rol'
  ]);

  participantes.forEach(function(p) {
    if (p.rol === 'Capitán' && !esCompetidorCapitan(d, p)) return;
    const dorsal = siguienteDorsal(h);
    h.appendRow([id, dorsal, p.nombre, p.sexo || '', d.categoria, p.rol]);
  });
}

function esCompetidorCapitan(d, p) {
  return Boolean(d.capitan && d.capitan.compite === true);
}

function siguienteDorsal(h) {
  const cantidad = Math.max(0, h.getLastRow() - 1) + 1;
  return Utilities.formatString('%03d', cantidad);
}

function solicitaHospedaje(p) {
  return String(p.hospedaje || '').toLowerCase() === 'organizacion';
}

function necesitaErgo(c) {
  return String(c.nacionalidad || '').toLowerCase() === 'extranjera' && Number(c.edad) > 35;
}

function generarWord(id, d, participantes, carpeta) {
  const doc = DocumentApp.create('Inscripción ' + id + ' — ' + d.capitan.nombre);
  const cuerpo = doc.getBody();

  if (LOGO_FILE_ID) {
    try {
      const logo = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
      const pLogo = cuerpo.appendParagraph('');
      pLogo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      const img = pLogo.appendInlineImage(logo);
      const alto = 90;
      img.setHeight(alto).setWidth(Math.round(img.getWidth() * alto / img.getHeight()));
    } catch (e) {
      console.warn('No se pudo insertar el logo: ' + e);
    }
  }

  const titulo = cuerpo.appendParagraph('COPA INTEGRACIÓN DNB 2026');
  titulo.setHeading(DocumentApp.ParagraphHeading.TITLE);
  titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titulo.editAsText().setForegroundColor('#A61E24').setBold(true);

  const sub = cuerpo.appendParagraph('FICHA DE INSCRIPCIÓN CONFIRMADA');
  sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  sub.editAsText().setBold(true).setForegroundColor('#C49A42');

  tabla(cuerpo, [
    ['N.º inscripción', id],
    ['Fecha de inscripción', formatearFecha(new Date(), "dd/MM/yyyy HH:mm 'hs'")],
    ['Modalidad seleccionada', d.categoria]
  ]);

  encabezado(cuerpo, 'DATOS DEL CAPITÁN');
  tabla(cuerpo, [
    ['Nombre y apellido', d.capitan.nombre],
    ['Cédula de identidad', d.capitan.ci || ''],
    ['Teléfono', d.capitan.celular || ''],
    ['Correo electrónico', d.capitan.email || ''],
    ['Hospedaje', solicitaHospedaje(d.capitan)
      ? 'SOLICITA HOSPEDAJE DE LA ORGANIZACIÓN'
      : 'NO SOLICITA HOSPEDAJE']
  ]);

  encabezado(cuerpo, 'COMPETIDORES');
  const filas = [['N.º', 'Nombre', 'C.I./Documento', 'Sexo', 'Edad', 'Modalidad', 'Hospedaje', 'Carnet', 'Ergometría']];
  participantes.filter(function(p) { return p.rol === 'Competidor'; }).forEach(function(p) {
    filas.push([
      String(p.numero),
      p.nombre,
      p.documento || '',
      p.sexo || '',
      String(p.edad || ''),
      d.categoria,
      solicitaHospedaje(p) ? 'Solicita organización' : 'No solicita',
      p.linkCarnet ? 'PRESENTADO' : 'FALTA',
      p.linkErgo ? 'PRESENTADA' : (necesitaErgo(p) ? 'FALTA' : 'No requiere')
    ]);
  });
  estilizarTabla(cuerpo.appendTable(filas), true);

  encabezado(cuerpo, 'RESUMEN DE HOSPEDAJE');
  tabla(cuerpo, [
    ['Personas registradas', String(participantes.length)],
    ['Solicitan hospedaje de la organización', String(participantes.filter(solicitaHospedaje).length)],
    ['No solicitan hospedaje', String(participantes.filter(function(p) { return !solicitaHospedaje(p); }).length)],
    ['Total de camas necesarias', String(participantes.filter(solicitaHospedaje).length)]
  ]);

  cuerpo.appendParagraph('');
  cuerpo.appendParagraph('Carpeta de documentación: ' + carpeta.getUrl())
    .editAsText().setForegroundColor('#1155CC').setUnderline(true);

  doc.saveAndClose();

  const url = 'https://docs.google.com/document/d/' + doc.getId() + '/export?format=docx';
  const blobDocx = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName('Ficha_' + id + '.docx');

  const archivoDocx = carpeta.createFile(blobDocx);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return archivoDocx;
}

function encabezado(cuerpo, texto) {
  const h = cuerpo.appendParagraph(texto);
  h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  h.editAsText().setForegroundColor('#A61E24').setBold(true);
}

function tabla(cuerpo, filas) {
  const t = cuerpo.appendTable(filas);
  estilizarTabla(t, false);
  for (let i = 0; i < t.getNumRows(); i++) {
    t.getRow(i).getCell(0).editAsText().setBold(true);
  }
  return t;
}

function estilizarTabla(t, conEncabezado) {
  t.setBorderColor('#BBBBBB');
  if (conEncabezado && t.getNumRows()) {
    const fila = t.getRow(0);
    for (let c = 0; c < fila.getNumCells(); c++) {
      fila.getCell(c).setBackgroundColor('#A61E24')
        .editAsText().setForegroundColor('#FFFFFF').setBold(true);
    }
  }
}

function enviarConfirmacion(id, d, participantes) {
  if (!d.capitan.email) return;

  const solicita = participantes.filter(solicitaHospedaje).length;
  const cuerpo =
    'Hola ' + d.capitan.nombre + ',\n\n' +
    'Confirmamos la inscripción a la COPA INTEGRACIÓN DNB 2026.\n\n' +
    'N.º de inscripción: ' + id + '\n' +
    'Modalidad seleccionada: ' + d.categoria + '\n' +
    'Competidores: ' + (d.competidores || []).length + '\n' +
    'Personas que solicitan hospedaje: ' + solicita + '\n\n' +
    'La inscripción quedó registrada correctamente.\n\n' +
    'Dirección Nacional de Bomberos';

  MailApp.sendEmail({
    to: d.capitan.email,
    subject: 'Inscripción ' + id + ' confirmada — Copa Integración DNB 2026',
    body: cuerpo
  });
}

function enviarWordOrganizacion(id, d, archivoDocx, carpeta) {
  if (!MAIL_ORGANIZACION) return;

  MailApp.sendEmail({
    to: MAIL_ORGANIZACION,
    subject: '[Copa DNB 2026] ' + id + ' — ' + d.categoria + ' — ' + d.capitan.nombre,
    body:
      'Nueva inscripción registrada.\n\n' +
      'Capitán: ' + d.capitan.nombre + '\n' +
      'Modalidad seleccionada: ' + d.categoria + '\n' +
      'Competidores: ' + (d.competidores || []).length + '\n\n' +
      'Carpeta de documentación: ' + carpeta.getUrl(),
    attachments: [archivoDocx.getBlob()]
  });
}

function formatearFecha(fecha, formato) {
  return Utilities.formatDate(fecha, 'America/Montevideo', formato);
}

function respuesta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
