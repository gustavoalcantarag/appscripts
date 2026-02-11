function enviarCorreosDesdeSheet() {
  var ss = SpreadsheetApp.openById('10Vt7Yr9xssqWpwLoGvJS4FJeMzz-LeahT1jP-h5QhsE');
  var sheet = ss.getSheetByName('Html');
  var datos = sheet.getDataRange().getValues();
  // Leer la fecha/hora de la celda A2 de la hoja Hora
  var sheetHora = ss.getSheetByName('Hora');
  var fechaHoraLimite = sheetHora.getRange('A2').getValue();

  // 1. Eliminar filas con 'Enviado' en la columna G (desde la segunda fila)
  //    o cuya fecha en columna A (Html) sea menor que la fecha/hora de la hoja Hora (A2)
  for (var i = datos.length - 1; i >= 1; i--) {
    var estado = datos[i][6]; // Columna G
    var fechaHtml = datos[i][0]; // Columna A
    var eliminarPorEstado = estado && estado.toString().trim().toLowerCase() === 'enviado';
    var eliminarPorFecha = false;
    // Validar si fechaHtml es menor que fechaHoraLimite
    if (fechaHtml && fechaHoraLimite) {
      // Convertir a Date si es necesario
      var fechaHtmlDate = (fechaHtml instanceof Date) ? fechaHtml : new Date(fechaHtml);
      var fechaHoraLimiteDate = (fechaHoraLimite instanceof Date) ? fechaHoraLimite : new Date(fechaHoraLimite);
      eliminarPorFecha = fechaHtmlDate < fechaHoraLimiteDate;
    }
    if (eliminarPorEstado || eliminarPorFecha) {
      sheet.deleteRow(i + 1);
    }
  }

  // Volver a cargar los datos después de eliminar filas
  datos = sheet.getDataRange().getValues();

  // 2. Enviar correos para las filas pendientes
  for (var i = 1; i < datos.length; i++) {
    var estado = datos[i][6]; // Columna G
    var asunto = datos[i][1]; // Columna B
    var remitente = datos[i][2]; // Columna C
    var destinatario = 'gusnaty76@gmail.com'; // para el envío
    var destinatarioReal = datos[i][3]; // Columna D
    var ccReal = datos[i][4]; // Columna E
    var htmlBody = datos[i][5]; // Columna F

    if (!estado || estado.toString().trim().toLowerCase() !== 'enviado') {
      if (htmlBody && htmlBody.trim() !== '') {
        var cuerpoCorreo = '';
        cuerpoCorreo += '<b>De:</b> ' + remitente + '<br>';
        cuerpoCorreo += '<b>Para:</b> ' + destinatarioReal + '<br>';
        cuerpoCorreo += (ccReal ? ('<b>CC:</b> ' + ccReal + '<br>') : '');
        cuerpoCorreo += '<hr>' + htmlBody;

        GmailApp.sendEmail(
          destinatario,
          asunto,
          '',
          {
            htmlBody: cuerpoCorreo,
            name: 'Reenvío automático'
          }
        );
        // 3. Marcar como 'Enviado' en la columna G
        sheet.getRange(i + 1, 7).setValue('Enviado');
      }
    }
  }
}
