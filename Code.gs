/**
 * SOY YOGA — Backend de reservas (Google Apps Script)
 * - Guarda cada reserva en la hoja "Reservas"
 * - Controla cupos: Aeroyoga 3 por día, demás clases 5 por clase
 * - Guarda comprobantes de transferencia en Google Drive
 * - Te envía un correo por cada reserva y otro de confirmación al alumno
 *
 * Para liberar un lugar: cambia la columna "Estado" a "Cancelada" (o borra la fila).
 */

const OWNER_EMAIL = 'itziacasaos@gmail.com';
const SHEET_NAME = 'Reservas';
const FOLDER_NAME = 'Soy Yoga - Comprobantes';
const TZ = 'America/Mexico_City';
const CUPO_AEROYOGA = 3;
const CUPO_REGULAR = 5;

const HEADERS = ['Registrado', 'Estado', 'Clase', 'Fecha', 'Hora', 'Nombre', 'WhatsApp', 'Correo',
  'Experiencia', 'Notas / lesiones', 'Método de pago', 'Monto', 'Comprobante', 'Clave'];
const COL_ESTADO = 1, COL_CLAVE = 13;

/** Ejecuta esta función UNA vez desde el editor para dar permisos y crear la hoja. */
function setup() {
  getSheet_();
  getFolder_();
  Logger.log('Listo. Ahora publica: Implementar > Nueva implementación > Aplicación web.');
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'availability') return json_({ ok: true, counts: getCounts_() });
  return json_({ ok: true, message: 'Soy Yoga API activa' });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, message: 'Datos inválidos' }); }

  // Validaciones básicas
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data.fecha) || !/^\d{2}:\d{2}$/.test(data.hora))
    return json_({ ok: false, message: 'Fecha u hora inválida' });
  if (!data.nombre || String(data.telefono || '').replace(/\D/g, '').length < 10)
    return json_({ ok: false, message: 'Faltan nombre o WhatsApp' });
  if (data.fecha < today_()) return json_({ ok: false, message: 'Esa fecha ya pasó' });
  if (data.metodo === 'transferencia' && !(data.comprobante && data.comprobante.data))
    return json_({ ok: false, message: 'Falta el comprobante de transferencia' });

  const clave = data.fecha + '|' + data.hora;
  const cupo = capacityFor_(data.fecha, data.hora);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let fileUrl = '', blob = null;
  try {
    const counts = getCounts_();
    if ((counts[clave] || 0) >= cupo) return json_({ ok: false, error: 'full' });

    if (data.comprobante && data.comprobante.data) {
      const c = data.comprobante;
      blob = Utilities.newBlob(Utilities.base64Decode(c.data), c.mime || 'image/jpeg',
        data.fecha + ' ' + data.nombre + ' - ' + (c.name || 'comprobante'));
      const file = getFolder_().createFile(blob);
      fileUrl = file.getUrl();
    }

    const estado = data.metodo === 'mercadopago'
      ? 'Pendiente: verificar pago Mercado Pago'
      : 'Comprobante recibido: verificar transferencia';

    getSheet_().appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
      estado,
      data.clase,
      data.fechaTexto || data.fecha,
      "'" + (data.horaTexto || data.hora),
      data.nombre,
      "'" + data.telefono,
      data.email || '',
      data.experiencia || '',
      data.notas || '',
      data.metodo === 'mercadopago' ? 'Mercado Pago' : 'Transferencia',
      data.monto,
      fileUrl,
      clave
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try { sendEmails_(data, fileUrl, blob); } catch (err) { Logger.log(err); }
  return json_({ ok: true });
}

/* ---------- helpers ---------- */

function capacityFor_(fecha, hora) {
  const dow = new Date(fecha + 'T12:00:00').getDay();
  return ((dow === 0 || dow === 6) && hora === '09:30') ? CUPO_AEROYOGA : CUPO_REGULAR;
}

function getCounts_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const counts = {};
  if (last < 2) return counts;
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const hoy = today_();
  rows.forEach(r => {
    const estado = String(r[COL_ESTADO]);
    const clave = String(r[COL_CLAVE]);
    if (!clave || /cancel/i.test(estado)) return;
    if (clave.split('|')[0] < hoy) return;
    counts[clave] = (counts[clave] || 0) + 1;
  });
  return counts;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#2B1D2C').setFontColor('#F3ECE4');
  }
  return sh;
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function sendEmails_(d, fileUrl, blob) {
  const mp = d.metodo === 'mercadopago';
  const wa = 'https://wa.me/52' + String(d.telefono).replace(/\D/g, '').slice(-10);
  const rows = [
    ['Clase', d.clase], ['Fecha', d.fechaTexto], ['Hora', d.horaTexto],
    ['Nombre', d.nombre], ['WhatsApp', '<a href="' + wa + '">' + esc_(d.telefono) + '</a>'],
    ['Correo', d.email || '—'], ['Experiencia', d.experiencia], ['Notas', d.notas || '—'],
    ['Pago', (mp ? 'Mercado Pago (verifica en tu cuenta)' : 'Transferencia') + ' · $' + d.monto],
    ['Comprobante', fileUrl ? '<a href="' + fileUrl + '">Ver en Drive</a> (también adjunto)' : '—']
  ];
  const html = '<div style="font-family:Arial,sans-serif;color:#2B1D2C">' +
    '<h2 style="color:#B0452A;margin:0 0 12px">Nueva reserva · clase muestra</h2>' +
    '<table cellpadding="6" style="border-collapse:collapse">' +
    rows.map(r => '<tr><td style="color:#6E5F72"><b>' + r[0] + '</b></td><td>' +
      (r[0] === 'WhatsApp' || r[0] === 'Comprobante' ? r[1] : esc_(r[1])) + '</td></tr>').join('') +
    '</table></div>';

  const opts = { htmlBody: html, name: 'Soy Yoga Reservas' };
  if (d.email) opts.replyTo = d.email;
  if (blob) opts.attachments = [blob];
  MailApp.sendEmail(OWNER_EMAIL, 'Nueva reserva: ' + d.clase + ' · ' + d.fechaTexto + ' ' + d.horaTexto + ' · ' + d.nombre, '', opts);

  if (d.email) {
    const body = '<div style="font-family:Arial,sans-serif;color:#2B1D2C">' +
      '<h2 style="color:#B0452A">¡Gracias, ' + esc_(d.nombre.split(' ')[0]) + '!</h2>' +
      '<p>Recibimos tu reserva para tu clase muestra en <b>Soy Yoga</b>:</p>' +
      '<p><b>' + esc_(d.clase) + '</b><br>' + esc_(d.fechaTexto) + ', ' + esc_(d.horaTexto) + '</p>' +
      (mp ? '<p>Si aún no completas tu pago de $' + d.monto + ' por Mercado Pago, hazlo aquí: <a href="https://mpago.li/2vNPoHg">pagar clase muestra</a>.</p>'
          : '<p>Recibimos tu comprobante de transferencia. Te confirmaremos por WhatsApp.</p>') +
      '<p>Llega 10 minutos antes y trae ropa cómoda. ¡Nos da mucho gusto que seas parte de la comunidad! 🙏</p></div>';
    MailApp.sendEmail(d.email, 'Tu clase muestra en Soy Yoga · ' + d.fechaTexto, '', { htmlBody: body, name: 'Soy Yoga', replyTo: OWNER_EMAIL });
  }
}
