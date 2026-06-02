const CONFIG = {
  SHEET_ID: '1eJ6zw2HtJ-0MIOOhx9G5QgP5X1IyGBAtwwOKXTDoNiQ',
  SHEET_NAME: 'Demos',
  DRIVE_FOLDER_ID: '10hhMIGUswJ8xJgsUXdfqhioiMAK8-1Zu',
  CORREO_EMPRESA: 'tntone.audio@gmail.com',
  TNTONE_WHATSAPP: '59170514745',
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const driveLink = subirArchivo(data);
    guardarEnSheets(data, driveLink);
    enviarEmailEmpresa(data, driveLink);
    return respuestaOk({ status: 'ok', message: 'Demo registrada correctamente.' });
  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    return respuestaOk({ status: 'error', message: err.message });
  }
}

function doGet() {
  return respuestaOk({ status: 'ok', message: 'TNTone API activa.' });
}

function subirArchivo(data) {
  if (!data.archivoBase64 || !data.archivoNombre) return '(sin archivo)';
  const bytes   = Utilities.base64Decode(data.archivoBase64);
  const blob    = Utilities.newBlob(bytes, data.archivoMime, data.archivoNombre);
  const folder  = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const archivo = folder.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return archivo.getUrl();
}

function guardarEnSheets(data, driveLink) {
  const hoja = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);

  if (hoja.getLastRow() === 0) {
    hoja.appendRow([
      'Fecha', 'Artista / Banda', 'Nombre Demo', 'Tipo de Servicio',
      'Qué busca', 'Correo', 'Código País', 'Teléfono',
      'Archivo (Drive)', 'Estado', 'Link WhatsApp'
    ]);
    const cabecera = hoja.getRange(1, 1, 1, 11);
    cabecera.setBackground('#0a0a0a');
    cabecera.setFontColor('#C9A84C');
    cabecera.setFontWeight('bold');
    hoja.setFrozenRows(1);
  }

  const fila = hoja.getLastRow() + 1;

  hoja.appendRow([
    data.fecha || new Date().toLocaleString('es-BO'),
    data.artista,
    data.demo,
    data.servicio,
    data.busca,
    data.email,
    data.codigoPais,
    data.telefono,
    driveLink,
    'Pendiente',
    ''
  ]);

  const celdaEstado = hoja.getRange(fila, 10);
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'Aprobado', 'Rechazado'], true)
    .build();
  celdaEstado.setDataValidation(regla);
  hoja.getRange(fila, 1, 1, 11).setBackground(fila % 2 === 0 ? '#141414' : '#0f0f0f');
}

function enviarEmailEmpresa(data, driveLink) {
  const asunto = `[TNTone] Nueva demo recibida: ${data.artista} — ${data.demo}`;

  const cuerpoHTML = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <style>
      body { background:#0a0a0a; margin:0; padding:0; font-family:'Georgia',serif; }
      .wrapper { max-width:580px; margin:0 auto; background:#111; border:1px solid #8B6914; }
      .header { background:linear-gradient(135deg,#0a0a0a,#1a1400); padding:36px 40px; text-align:center; border-bottom:2px solid #C9A84C; }
      .logo { font-size:2rem; letter-spacing:8px; color:#C9A84C; font-weight:bold; margin:0; }
      .tagline { color:#8B6914; font-size:0.75rem; letter-spacing:4px; margin-top:6px; text-transform:uppercase; }
      .body { padding:36px 40px; }
      .title { color:#C9A84C; font-size:1.1rem; letter-spacing:2px; text-transform:uppercase; margin:0 0 24px; border-bottom:1px solid rgba(201,168,76,0.25); padding-bottom:14px; }
      .row { display:flex; justify-content:space-between; padding:10px 0; border-bottom:1px solid rgba(201,168,76,0.1); }
      .lbl { color:#8B6914; font-size:0.8rem; letter-spacing:2px; text-transform:uppercase; }
      .val { color:#e8dfc8; font-size:0.95rem; text-align:right; max-width:60%; word-break:break-word; }
      .link { color:#C9A84C; text-decoration:none; }
      .footer { padding:20px 40px; text-align:center; background:#0a0a0a; border-top:1px solid rgba(201,168,76,0.2); }
      .ft { color:#4a3a00; font-size:0.75rem; letter-spacing:2px; }
    </style>
  </head>
  <body>
    <div class="wrapper">
      <div class="header">
        <p class="logo">TNTone</p>
        <p class="tagline">Nueva Demo Recibida</p>
      </div>
      <div class="body">
        <p class="title">Datos del Proyecto</p>
        <div class="row"><span class="lbl">Artista</span><span class="val">${data.artista}</span></div>
        <div class="row"><span class="lbl">Demo</span><span class="val">${data.demo}</span></div>
        <div class="row"><span class="lbl">Servicio</span><span class="val">${data.servicio}</span></div>
        <div class="row"><span class="lbl">Busca</span><span class="val">${data.busca}</span></div>
        <div class="row"><span class="lbl">Correo</span><span class="val"><a href="mailto:${data.email}" class="link">${data.email}</a></span></div>
        <div class="row"><span class="lbl">WhatsApp</span><span class="val">${data.codigoPais} ${data.telefono}</span></div>
        <div class="row"><span class="lbl">Fecha</span><span class="val">${data.fecha || new Date().toLocaleString()}</span></div>
        <div class="row" style="border:none"><span class="lbl">Archivo</span><span class="val"><a href="${driveLink}" class="link">Ver en Drive ↗</a></span></div>
      </div>
      <div class="footer">
        <p class="ft">© TNTone · Ingeniería de Sonido</p>
      </div>
    </div>
  </body>
  </html>`;

  MailApp.sendEmail({
    to: CONFIG.CORREO_EMPRESA,
    subject: asunto,
    htmlBody: cuerpoHTML
  });
}

function onEditHoja(e) {
  const hoja = e.source.getSheetByName(CONFIG.SHEET_NAME);
  if (!hoja) return;

  const rango = e.range;
  if (rango.getColumn() !== 10 || rango.getRow() < 2) return;

  const fila   = rango.getRow();
  const estado = rango.getValue();

  if (estado !== 'Aprobado' && estado !== 'Rechazado') return;

  const artista  = hoja.getRange(fila, 2).getValue();
  const servicio = hoja.getRange(fila, 4).getValue();
  const codigo   = hoja.getRange(fila, 7).getValue();
  const telefono = hoja.getRange(fila, 8).getValue();
  const numero   = (codigo + telefono).replace(/\D/g, '');

  let mensaje = '';

  if (estado === 'Aprobado') {
    mensaje =
      `¡Hola, ${artista}! 🎶 Somos TNTone.\n\n` +
      `Tu demo fue escuchada y nos dejó sin palabras — quedamos realmente impresionados con lo que traes. ` +
      `Con mucho gusto te confirmamos que tu proyecto ha sido *aprobado* para *${servicio}*. 🏆\n\n` +
      `Para que podamos comenzar a trabajar en él, necesitamos que realices el pago del *50% del valor del servicio* mediante nuestro código QR. ` +
      `Una vez confirmado el pago, te indicaremos la fecha estimada de entrega del proyecto finalizado. ✨\n\n` +
      `No dejes de innovar — el mundo necesita escucharte. 🎚️\n\n` +
      `_Equipo TNTone_`;
  } else {
    mensaje =
      `¡Hola, ${artista}! Somos TNTone. 🎶\n\n` +
      `Queremos agradecerte sinceramente por confiar en nosotros y enviarnos tu trabajo. ` +
      `Tu demo fue escuchada con mucha atención y en verdad el talento está ahí.\n\n` +
      `Lamentablemente, en este momento nuestra agenda no nos permite tomar nuevos proyectos y no queremos comprometernos sin poder darte el tiempo y la dedicación que tu música merece.\n\n` +
      `Pero no te pierdas de vista: estaremos atentos y nos comunicaremos contigo en cuanto tengamos disponibilidad. Los proyectos con potencial no se olvidan. 🙌\n\n` +
      `_Equipo TNTone_`;
  }

  const link = `https://wa.me/${numero}?text=${encodeURIComponent(mensaje)}`;

  const celdaEstado = hoja.getRange(fila, 10);
  if (estado === 'Aprobado') {
    celdaEstado.setBackground('#1a3a2a').setFontColor('#5ec99a');
  } else {
    celdaEstado.setBackground('#3a1a1a').setFontColor('#e07070');
  }

  const celdaLink = hoja.getRange(fila, 11);
  celdaLink.setFormula(`=HYPERLINK("${link}","Enviar WhatsApp → ${estado}")`);
  celdaLink.setFontColor('#C9A84C');
}

function respuestaOk(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
