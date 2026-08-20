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

// Fuerza un valor a texto literal anteponiendo un apóstrofe — el mismo
// truco que usa Sheets cuando lo tipeás a mano. Sin esto, un WhatsApp como
// "+54 9..." se interpreta como el inicio de una fórmula y queda #ERROR!.
// setNumberFormat('@') NO alcanza para evitar esto: esa detección pasa al
// momento de escribir el valor, no depende del formato de la celda.
function comoTexto(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  return "'" + String(valor);
}

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const hoja = obtenerHoja();
    const valores = [
      comoTexto(datos.nombre),
      comoTexto(datos.whatsapp),
      comoTexto(datos.email),
      comoTexto(datos.negocio),
      comoTexto(datos.actividad),
      comoTexto(datos.cantidadTurnos),
      comoTexto(datos.cantidadProfesionales),
      comoTexto(Array.isArray(datos.sistemaActual) ? datos.sistemaActual.join(', ') : ''),
      comoTexto(datos.tiempoDedicado),
      comoTexto(Array.isArray(datos.problemas) ? datos.problemas.join(', ') : ''),
      comoTexto(datos.resultadoCalculado),
      comoTexto(Array.isArray(datos.recomendaciones) ? datos.recomendaciones.join(', ') : ''),
      comoTexto(datos.fecha)
    ];

    // Escritura manual (no appendRow): se formatean las celdas de texto
    // como Plain Text ANTES de escribirlas, en la fila exacta, con
    // setValues — mismo patrón que ya usás para fecha/hora en tu app de
    // turnos. La columna A queda aparte porque es una fecha real (Date).
    const fila = hoja.getLastRow() + 1;
    hoja.getRange(fila, 1).setValue(new Date());
    const rangoTexto = hoja.getRange(fila, 2, 1, valores.length);
    rangoTexto.setNumberFormat('@');
    rangoTexto.setValues([valores]);

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
