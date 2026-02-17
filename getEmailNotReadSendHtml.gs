function reenviarCorreos() {
  var sheetId = '10Vt7Yr9xssqWpwLoGvJS4FJeMzz-LeahT1jP-h5QhsE'; // Mismo ID de la hoja compartida
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('CorreosProcesados');
  var data = sheet.getDataRange().getValues();
  var filasAEliminar = [];

  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    var asunto = data[i][2];
    var html = data[i][3];
    var remitente = data[i][4];
    var destinatario = data[i][5];
    var ccOriginal = data[i][6];
    var bccOriginal = data[i][7];

    MailApp.sendEmail({
      to: 'gusnaty76@gmail.com',
      subject: '[Reenvío] ' + asunto,
      htmlBody:
        '<b>Remitente original:</b> ' + remitente + '<br>' +
        '<b>Destinatario original:</b> ' + destinatario + '<br>' +
        '<b>CC original:</b> ' + ccOriginal + '<br>' +
        '<b>BCC original:</b> ' + bccOriginal + '<br><br>' +
        html
    });
    filasAEliminar.push(i+1); // Guarda el número de fila para eliminar
  }

  // Elimina las filas de abajo hacia arriba para evitar problemas de indexación
  for (var j = filasAEliminar.length - 1; j >= 0; j--) {
    sheet.deleteRow(filasAEliminar[j]);
  }
}
