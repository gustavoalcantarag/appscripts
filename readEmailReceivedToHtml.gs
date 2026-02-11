function correoToHtml() {
  var sheetHora = SpreadsheetApp.openById('10Vt7Yr9xssqWpwLoGvJS4FJeMzz-LeahT1jP-h5QhsE').getSheetByName('Hora');
  var sheetHtml = SpreadsheetApp.openById('10Vt7Yr9xssqWpwLoGvJS4FJeMzz-LeahT1jP-h5QhsE').getSheetByName('Html');
  var ultimaFecha = sheetHora.getRange('A2').getValue();
  if (!ultimaFecha) ultimaFecha = new Date(0);

  var threads = GmailApp.search('is:unread');
  var nuevaFecha = new Date(ultimaFecha);

  var registros = [];

  for (var i = 0; i < threads.length; i++) {
    var mensajes = threads[i].getMessages();
    for (var j = 0; j < mensajes.length; j++) {
      var mensaje = mensajes[j];
      if (!mensaje.isUnread()) continue;
      var fechaCorreo = mensaje.getDate();
      if (fechaCorreo > ultimaFecha) {
        var fechaHora = Utilities.formatDate(fechaCorreo, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        var asunto = mensaje.getSubject();
        var remitente = mensaje.getFrom();
        var destinatario = mensaje.getTo();
        var cc = mensaje.getCc();
        var htmlBody = mensaje.getBody();
        registros.push([fechaHora, asunto, remitente, destinatario, cc, htmlBody]);
        if (fechaCorreo > nuevaFecha) {
          nuevaFecha = fechaCorreo;
        }
      }
    }
  }

  // Escribir FechaHora, Asunto, Remitente, Destinatario, CC y HTML en la hoja 'Html' desde la segunda fila
  if (registros.length > 0) {
    var ultimaFila = sheetHtml.getLastRow();
    sheetHtml.getRange(ultimaFila + 1, 1, registros.length, 6).setValues(registros);
  }

  // Actualizar la celda A2 de la hoja 'Hora' con la fecha más reciente procesada
  sheetHora.getRange('A2').setValue(nuevaFecha);
}
