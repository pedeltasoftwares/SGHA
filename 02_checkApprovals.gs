/*
VERIFICA SI LA SOLCITUD DE HORAS ADICIONALES FUE APROBADA
*/
function checkApprovals() {

  // Obtener los correos de la etiqueta especificada
  var labelFolder = GmailApp.getUserLabelByName("SGHA")

  // Obtiene todos los hilos (threads) dentro de la carpeta de la etiqueta
  var emails = labelFolder.getThreads();

  // Itera por los correos
  for (var i = 0; i < emails.length ; i ++){

    //Mensajes asociados a cada correo
    var mensajes = emails[i].getMessages()

    //Si hay mas de un mensaje en el mismo hilo se asume que la hora adicional fue aprobada
    if (mensajes.length > 1)
    {
      //Obtiene los metadatos del correo automatico enviado
      var metadata = getEmailMetadata(mensajes[0])

      //Obtiene los datos de la solicitud aprobada
      var array_solicitud = getApprovalData(metadata.Body)

      //Escribe "Si" en la casilla de aprobación y envia el correo al colaborador
      writeApproval(array_solicitud)

    }
    
  }
  
}

/*
/Obtiene los datos de la solicitud aprobada
*/
function getApprovalData(body){

  //Inicializa el array
  var approvalData = []

  //Encuentra el nombre del colaborador
  var regex = /COLABORADOR: (.+)/;
  var match = regex.exec(body);
  approvalData.push(match[1].trim())

  //Encuentra el email del colaborador
  var regex = /EMAIL: (.+)/;
  var match = regex.exec(body);
  approvalData.push(match[1].trim())

  //Encuentra la fecha del cumplimiento de la hora adicional
  var regex = /FECHA: (.+)/;
  var match = regex.exec(body);
  approvalData.push(match[1].trim())

  //Encuentra el area
  var regex = /ÁREA: (.+)/;
  var match = regex.exec(body);
  approvalData.push(match[1].trim())

  //Encuentra el proyecto
  var regex = /PROYECTO: ([^\n]+)/;
  var match = regex.exec(body);
  approvalData.push(match[1].trim())

  //Encuentra la descripción de la labor
  var regex = /DESCRIPCIÓN LABOR: (.+)/;
  var match = regex.exec(body);
  approvalData.push(match[1].trim())

  return approvalData

}

/*
CONFIRMA APROBACIÓN DE HORAS ADICIONALES EN EL REGISTRO ADECUADO
*/
function writeApproval(array_solicitud)
{

  //Obtiene el libro actual
  var book = SpreadsheetApp.getActiveSpreadsheet();

  //Obtiene la hoja de la dependencia
  var sheetArea = book.getSheetByName(array_solicitud[3]);

  //Obtiene el registro
  var registros =  sheetArea.getRange("A2:F").getDisplayValues();

  //Filtra los arrays que contengan registros
  var registros = filterBlankSpacesSubarray(registros)

  //Itera por los registros
  for (var i = 0; i < registros.length ; i++)
  { 
    //Estado de aprobación
    var estadoAprobacion = sheetArea.getRange("I"+(i+2)).getDisplayValue();

    //Encuentra la línea de registro que coincide con la solicitud aprobada
    if(registros[i][0] === array_solicitud[0] && registros[i][1] === array_solicitud[1] && registros[i][2] === array_solicitud[2] && registros[i][3] === array_solicitud[4] && registros[i][4] === array_solicitud[5] && estadoAprobacion !== "Sí")
    {
      //Escribe "si" en la casilla de aprobación
      sheetArea.getRange("I"+(i+2)).setValue("Sí")

      //Envía el correo de notificación de aprobación
      sendEmailEmployee(sheetArea,i,array_solicitud)

    }
  }
}