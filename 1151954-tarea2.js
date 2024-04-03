// Función que se ejecuta al abrir la hoja de cálculo
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Crear un menú personalizado
    ui.createMenu('Combina Correspondencia')
      .addItem('Combinar y enviar', 'combinarYEnviar')
      .addToUi();
  }
  
  // Función para combinar correspondencia y enviar por correo electrónico
  function combinarYEnviar() {
    try {
      // Obtener la hoja de cálculo activa y la plantilla de documento
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = spreadsheet.getSheetByName('Hoja 1');
      var templateFileId = '1TQ10mnixcAnFHpQvh1u6U-AqEB6jcYC7bIB6pesHric';
  
      // Obtener los datos de la hoja de cálculo
      var dataRange = sheet.getDataRange();
      var data = dataRange.getValues();
      var headers = data[0];
      var numRows = data.length;
  
      // Recorrer los datos y combinar correspondencia
      for (var i = 1; i < numRows; i++) {
        var rowData = data[i];
  
        // Crear una copia de la plantilla y abrir el documento
        var copyFile = DriveApp.getFileById(templateFileId).makeCopy();
        var doc = DocumentApp.openById(copyFile.getId());
  
        // Reemplazar los marcadores de posición en la plantilla con datos de la hoja de cálculo
        for (var j = 0; j < headers.length; j++) {
          var placeholder = "{" + headers[j] + "}";
          var value = rowData[j];
          doc.getBody().replaceText(placeholder, value);
        }
  
        // Guardar los cambios en la plantilla
        doc.saveAndClose();
  
        // Obtener el correo electrónico del destinatario desde la hoja de cálculo
        var email = rowData[headers.indexOf('Email')]; // Suponiendo que el correo electrónico está en una columna llamada 'Email'
  
        // Enviar el correo utilizando Yet Another Mail Merge
        sendEmailWithYAMM(email, copyFile.getId());
  
        // Eliminar la copia del documento
        DriveApp.getFileById(copyFile.getId()).setTrashed(true);
      }
  
      // Mensaje de éxito
      SpreadsheetApp.getUi().alert('Se ha completado la combinación de correspondencia y el envío de correos electrónicos.');
    } catch (error) {
      // Manejar cualquier error que ocurra durante el proceso
      SpreadsheetApp.getUi().alert('Ha ocurrido un error: ' + error.message);
    }
  }
  
  // Función para enviar el correo utilizando Yet Another Mail Merge
  function sendEmailWithYAMM(recipient, fileId) {
    var subject = "Notas Actualizadas";
    var body = "Anexo pdf con sus notas actualizadas";
    var attachmentId = fileId;
    MailApp.sendEmail(recipient, subject, body, {attachments: [DriveApp.getFileById(attachmentId)]});
  }
  
