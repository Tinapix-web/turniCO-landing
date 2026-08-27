/**
 * TurniCO — Backend para el formulario de alta de clientes (/onboarding.html)
 * ==========================================================
 * Mismo patrón que google-apps-script-leads.js, con dos cosas más:
 * - Sube el logo (si lo adjuntaron) a una carpeta de Google Drive y
 *   guarda el link en la planilla.
 * - Te manda un mail avisando apenas alguien completa el formulario.
 *
 * CÓMO INSTALARLO
 * 1. Creá una Google Sheet NUEVA (ej: "TurniCO - Clientes"). Distinta a
 *    cualquier otra hoja que ya tengas (leads, turnos reales, etc).
 * 2. Extensiones > Apps Script, borrá todo lo que haya y pegá este archivo entero.
 * 3. Cambiá NOTIFICAR_A más abajo por tu email si hace falta (ya viene
 *    con valentinaaragonb@gmail.com puesto).
 * 4. Guardá (Ctrl+S).
 * 5. Implementar > Nueva implementación > tipo "Aplicación web".
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 *    La primera vez te va a pedir autorizar permisos (Sheets, Drive y Gmail) — es normal, aceptalo.
 * 6. Copiá la URL que te da (termina en /exec).
 * 7. Pegala en onboarding.html, dentro de CONFIG.clientesWebhookUrl = "...".
 *
 * Si más adelante cambiás el código: Implementar > Administrar
 * implementaciones > lápiz ✏️ > Nueva versión > Implementar. Si solo
 * guardás sin crear una versión nueva, el /exec sigue sirviendo el código viejo.
 *
 * TIP: además de este mail automático, podés activar el aviso nativo de
 * Google Sheets: en la hoja, Herramientas > Reglas de notificación >
 * "Cuando se agregue una fila" > "Notificarme de inmediato". Es un
 * respaldo extra, sin depender de este código.
 */

const NOMBRE_HOJA = 'Clientes';
const NOMBRE_CARPETA_LOGOS = 'TurniCO - Logos de clientes';
const NOTIFICAR_A = 'valentinaaragonb@gmail.com';

// Fuerza un valor a texto literal anteponiendo un apóstrofe — evita que
// Sheets interprete un WhatsApp con "+" como el inicio de una fórmula.
function comoTexto(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  return "'" + String(valor);
}

const NOMBRES_DIAS = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];

function formatearSucursales(sucursales) {
  if (!Array.isArray(sucursales)) return '';
  return sucursales.map(s => {
    const dias = Array.isArray(s.dias) ? s.dias.map(d => NOMBRES_DIAS[d] || d).join('/') : '';
    return `${s.nombre} — ${s.direccion}, ${s.barrio} | ${dias} | ${s.apertura}-${s.cierre} | turnos de ${s.duracion}`;
  }).join('\n');
}

function formatearServicios(servicios) {
  if (!Array.isArray(servicios)) return '';
  return servicios.map(s => `${s.nombre} (${s.duracion})${s.precio ? ' - ' + s.precio : ''}`).join('\n');
}

// Sube el logo (base64) a Drive y devuelve la URL pública, o '' si no vino logo.
function subirLogo(datos) {
  if (!datos.logoBase64) return '';
  try {
    const carpetas = DriveApp.getFoldersByName(NOMBRE_CARPETA_LOGOS);
    const carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(NOMBRE_CARPETA_LOGOS);
    const bytes = Utilities.base64Decode(datos.logoBase64);
    const nombreArchivo = `${datos.negocioNombre || 'logo'} - ${new Date().getTime()}`;
    const blob = Utilities.newBlob(bytes, datos.logoTipo || 'image/png', nombreArchivo);
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (err) {
    return 'ERROR AL SUBIR: ' + String(err);
  }
}

function enviarAvisoMail(datos, logoUrl) {
  try {
    const asunto = `🎉 Nuevo cliente TurniCO: ${datos.negocioNombre || '(sin nombre)'}`;
    const cuerpo = [
      `Nuevo formulario de alta completado.`,
      ``,
      `Negocio: ${datos.negocioNombre || '-'}`,
      `Rubro: ${datos.negocioRubro || '-'}`,
      `Plan: ${datos.plan || '-'}`,
      ``,
      `Contacto: ${datos.contactoNombre || '-'}`,
      `WhatsApp: ${datos.contactoWhatsapp || '-'}`,
      `Email: ${datos.contactoEmail || '-'}`,
      ``,
      `Instagram: ${datos.negocioInstagram || '-'}`,
      `Facebook: ${datos.negocioFacebook || '-'}`,
      `TikTok: ${datos.negocioTiktok || '-'}`,
      `Otra red: ${datos.negocioOtraRed || '-'}`,
      ``,
      `Sucursales:`,
      formatearSucursales(datos.sucursales) || '-',
      ``,
      `Servicios:`,
      formatearServicios(datos.servicios) || '-',
      ``,
      logoUrl ? `Logo: ${logoUrl}` : `Logo: no adjuntó (lo va a mandar por WhatsApp o email)`,
    ].join('\n');
    MailApp.sendEmail(NOTIFICAR_A, asunto, cuerpo);
  } catch (err) {
    // si falla el mail no queremos que se pierda el registro en la planilla
  }
}

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const logoUrl = subirLogo(datos);
    const hoja = obtenerHoja();
    hoja.appendRow([
      new Date(),
      comoTexto(datos.contactoNombre),
      comoTexto(datos.contactoWhatsapp),
      comoTexto(datos.contactoEmail),
      comoTexto(datos.negocioNombre),
      comoTexto(datos.negocioRubro),
      comoTexto(datos.negocioTagline),
      comoTexto(datos.negocioInstagram),
      comoTexto(datos.negocioFacebook),
      comoTexto(datos.negocioTiktok),
      comoTexto(datos.negocioOtraRed),
      comoTexto(datos.colorPrimario),
      comoTexto(datos.colorSecundario),
      comoTexto(logoUrl),
      comoTexto(formatearSucursales(datos.sucursales)),
      comoTexto(formatearServicios(datos.servicios)),
      comoTexto(datos.equipoCantidad),
      comoTexto(datos.equipoNombres),
      comoTexto(datos.plan),
      comoTexto(datos.fecha),
      // columnas JSON crudo, por si después querés reconstruir el theme.js con un script
      comoTexto(JSON.stringify(datos.sucursales || [])),
      comoTexto(JSON.stringify(datos.servicios || []))
    ]);

    enviarAvisoMail(datos, logoUrl);

    return ContentService.createTextOutput(JSON.stringify({ ok: true, logoUrl: logoUrl }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function obtenerHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA);
    hoja.appendRow([
      'Fecha registro', 'Nombre contacto', 'WhatsApp', 'Email',
      'Nombre negocio', 'Rubro', 'Eslogan', 'Instagram', 'Facebook', 'TikTok', 'Otra red',
      'Color principal', 'Color secundario', 'Logo (link)', 'Sucursales', 'Servicios',
      'Cant. equipo', 'Nombres equipo', 'Plan', 'Fecha formulario',
      'Sucursales (JSON)', 'Servicios (JSON)'
    ]);
  }
  return hoja;
}
