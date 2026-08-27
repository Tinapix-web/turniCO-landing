/**
 * TurniCO — Backend para el formulario de alta de clientes (/onboarding.html)
 * ==========================================================
 * Mismo patrón que google-apps-script-leads.js.
 *
 * CÓMO INSTALARLO
 * 1. Creá una Google Sheet NUEVA (ej: "TurniCO - Clientes"). Distinta a
 *    cualquier otra hoja que ya tengas (leads, turnos reales, etc).
 * 2. Extensiones > Apps Script, borrá todo lo que haya y pegá este archivo entero.
 * 3. Guardá (Ctrl+S).
 * 4. Implementar > Nueva implementación > tipo "Aplicación web".
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 * 5. Copiá la URL que te da (termina en /exec).
 * 6. Pegala en onboarding.html, dentro de CONFIG.clientesWebhookUrl = "...".
 *
 * Si más adelante cambiás el código, recordá: Implementar > Administrar
 * implementaciones > lápiz ✏️ > Nueva versión > Implementar. Si solo
 * guardás sin crear una versión nueva, el /exec sigue sirviendo el código viejo.
 */

const NOMBRE_HOJA = 'Clientes';

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

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
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
      comoTexto(datos.colorPrimario),
      comoTexto(datos.colorSecundario),
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
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
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
      'Nombre negocio', 'Rubro', 'Eslogan', 'Instagram', 'Facebook',
      'Color principal', 'Color secundario', 'Sucursales', 'Servicios',
      'Cant. equipo', 'Nombres equipo', 'Plan', 'Fecha formulario',
      'Sucursales (JSON)', 'Servicios (JSON)'
    ]);
  }
  return hoja;
}
