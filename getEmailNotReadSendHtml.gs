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

    var MAX_BODY = 190000; // margen seguro bajo el límite de 200 KB de Gmail
    if (html && html.length > MAX_BODY) {
      html = html.substring(0, MAX_BODY) + '<br><br><i>[Contenido truncado: el correo original superaba el límite de tamaño permitido]</i>';
    }

    var cuerpo = '<b>De:</b> ' + remitente + '<br>' +
      '<b>Para:</b> ' + destinatario + '<br>' +
      '<b>Cc:</b> ' + ccOriginal + '<br>' +
      html;

    Logger.log('Fila ' + (i+1) + ' | ID: ' + id + ' | Asunto: ' + asunto + ' | Tamaño del cuerpo: ' + cuerpo.length + ' caracteres');

    MailApp.sendEmail({
      to: 'gusnaty76@gmail.com',
      subject: asunto,
      htmlBody: cuerpo
    });
    filasAEliminar.push(i+1); // Guarda el número de fila para eliminar
  }

  // Elimina las filas de abajo hacia arriba para evitar problemas de indexación
  for (var j = filasAEliminar.length - 1; j >= 0; j--) {
    sheet.deleteRow(filasAEliminar[j]);
  }
}
