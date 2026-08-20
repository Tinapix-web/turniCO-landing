/**
 * TurniCO — Backend para los leads del Simulador de Turnos
 * ==========================================================
 * Mismo patrón que ya usás en tu app de turnos (Google Apps Script
 * como Web App gratuita + Google Sheets como base de datos).
 *
 * CÓMO INSTALARLO
 * 1. Creá una Google Sheet nueva (ej: "TurniCO - Leads").
 * 2. Extensiones > Apps Script, borrá lo que haya y pegá este archivo entero.
 * 3. Implementar > Nueva implementación > tipo "Aplicación web".
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 * 4. Copiá la URL que te da (termina en /exec).
 * 5. Pegala en index.html, dentro de CONFIG.leadsWebhookUrl = "...".
 *
 * A partir de ahí, cada vez que alguien complete el formulario del
 * simulador y toque "QUIERO QUE ME CONTACTEN", además de abrirse
 * WhatsApp, el lead completo queda guardado como una fila nueva acá.
 */

const NOMBRE_HOJA = 'Leads';

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const hoja = obtenerHoja();
    hoja.appendRow([
      new Date(),
      datos.nombre || '',
      datos.whatsapp || '',
      datos.email || '',
      datos.negocio || '',
      datos.actividad || '',
      datos.cantidadTurnos || '',
      datos.cantidadProfesionales || '',
      Array.isArray(datos.sistemaActual) ? datos.sistemaActual.join(', ') : '',
      datos.tiempoDedicado || '',
      Array.isArray(datos.problemas) ? datos.problemas.join(', ') : '',
      datos.resultadoCalculado || '',
      Array.isArray(datos.recomendaciones) ? datos.recomendaciones.join(', ') : '',
      datos.fecha || ''
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
      'Fecha registro', 'Nombre', 'WhatsApp', 'Email', 'Negocio', 'Actividad',
      'Cant. turnos', 'Cant. profesionales', 'Sistema actual', 'Tiempo dedicado',
      'Problemas', 'Resultado calculado', 'Recomendaciones', 'Fecha simulador'
    ]);
  }
  return hoja;
}
