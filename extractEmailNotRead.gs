function registrarCorreosNoLeidos() {
  var sheetId = '10Vt7Yr9xssqWpwLoGvJS4FJeMzz-LeahT1jP-h5QhsE'; // Pega aquí el ID de la hoja compartida
  var sheetCorreos = SpreadsheetApp.openById(sheetId).getSheetByName('CorreosProcesados');
  var sheetFecha = SpreadsheetApp.openById(sheetId).getSheetByName('UltimaFecha');
  var ultimaFecha = sheetFecha.getRange('A1').getValue();
  var threads = GmailApp.search('is:unread -is:muted');
  var nuevaFecha = new Date(ultimaFecha);

  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var id = message.getId();
      var fecha = message.getDate();
      var asunto = message.getSubject();
      var html = message.getBody();
      var remitente = message.getFrom();
      var destinatario = message.getTo();
      var cc = message.getCc();

      // Solo procesar si la fecha es posterior a la guardada
      if (fecha > nuevaFecha) {
        sheetCorreos.appendRow([id, fecha, asunto, html, remitente, destinatario, cc]);
        nuevaFecha = fecha;
      }
    });
  });

  // Actualiza la hoja UltimaFecha si se procesó algún correo nuevo
  if (nuevaFecha > new Date(ultimaFecha)) {
    sheetFecha.getRange('A1').setValue(nuevaFecha);
  }
}
