function sendEmailEmployee(sheetArea,i,array_solicitud){

  //Nombre del Lider
  var nombreLider = sheetArea.getRange("F"+(i+2)).getDisplayValue();

  //Asunto del correo
  var asuntoCorreo = "[P" + array_solicitud[4] + "] - Horas Adicionales Aprobadas";
  //Body
  var mensaje = "Estimado/a "+ array_solicitud[0] + ", \n\nLas horas adicionales que ha solicitado han sido aprobadas por " + nombreLider + ". A continuación, encontrará el detalle de su solicitud:\n\n"+
  "• FECHA: " + array_solicitud[2] + "\n" + 
  "• PROYECTO: " + array_solicitud[4] + "\n" + 
  "• DESCRIPCIÓN LABOR: " + array_solicitud[5] + "\n\nNo responder, esto es un mensaje automático.\n\n" +
  "Atentamente,\nSistema de Gestión de Horas Adicionales"

  GmailApp.sendEmail(array_solicitud[1], asuntoCorreo , mensaje, {from: "horas-adicionales-notificaciones@pedelta.com.co"});

}