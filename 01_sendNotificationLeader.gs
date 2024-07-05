/*
ENVÍA LA NOTIFICACIÓN DE SOLICITUD AL LIDER
*/
function sendNotificationLeader(book,dataLideres) {

  //Revisa los registros de las hojas "Infraestructura", "Puentes", "Edificaciones" para consultar cuales notificaciones no han sido enviadas
  var dependencias = ["Infraestructura","Puentes", "Edificaciones"];

  for (var dependencia of dependencias) {

    //Obtiene la hoja
    sheetArea = book.getSheetByName(dependencia);

    //Obtiene los valores de las notificaciones enviadas
    var notificacionEnviada = sheetArea.getRange("G1:G").getValues();

    //Filtra los arrays que contengan registros
    var notificacionEnviada = filterBlankSpacesSubarray(notificacionEnviada)

    //Envia la notificación siempre y cuando el valor de la casilla sea "No"
    for (var i = 0; i < notificacionEnviada.length ; i++)
    {
      if (notificacionEnviada[i][0] === "No")
      {
        //Envia el correo
        sendEmail(sheetArea,dataLideres,i,dependencia)
        //Cambia el estado
        sheetArea.getRange("G"+(i+1)).setValue("Sí")
        //Pone la hora de envío
        var currentDate = new Date();
        var hours = addLeadingZero(currentDate.getHours());
        var minutes = addLeadingZero(currentDate.getMinutes());
        var seconds = addLeadingZero(currentDate.getSeconds());
        var currentTime = hours + ":" + minutes + ":" + seconds;
        sheetArea.getRange("H"+(i+1)).setValue(currentDate)

      }
    }


  }
}

/*
ENVÍA EL CORREO
*/
function sendEmail(sheetArea,dataLideres,i,dependencia){

  //Nombre del Lider
  var nombreLider = sheetArea.getRange("F"+(i+1)).getDisplayValue();
  //Correo del lider
  var emailLider = leaderEmail(nombreLider,dataLideres)
  //Asunto del correo
  var asuntoCorreo = "Solicitud de Aprobación de Horas Adicionales - [" + sheetArea.getRange("A"+(i+1)).getDisplayValue()+"]";
  //Body
  var mensaje = "Estimado/a "+ nombreLider + ", \n\nEste es un mensaje automático para informarle que se ha recibido una nueva solicitud de aprobación de horas adicionales. A continuación, se detallan los datos de la solicitud:\n\n"+
  "• COLABORADOR: " + sheetArea.getRange("A"+(i+1)).getDisplayValue() + "\n" +
  "• EMAIL: " + sheetArea.getRange("B"+(i+1)).getDisplayValue() + "\n" +
  "• FECHA: " + sheetArea.getRange("C"+(i+1)).getDisplayValue() + "\n" + 
  "• ÁREA: " + dependencia + "\n" + 
  "• PROYECTO: " + sheetArea.getRange("D"+(i+1)).getDisplayValue() + "\n" + 
  "• DESCRIPCIÓN LABOR: " + sheetArea.getRange("E"+(i+1)).getDisplayValue() + "\n\n" +
  "Por favor, revise la solicitud y responda a este mismo correo para confirmar la aprobación de las horas adicionales. Una vez aprobadas, el colaborador será notificado.\n\n\nAtentamente,\nSistema de Gestión de Horas Adicionales"

  GmailApp.sendEmail(emailLider, asuntoCorreo , mensaje, {from: "horas-adicionales-notificaciones@pedelta.com.co"});

}

/*
ENCUENTRA EL CORREO DEL LIDER
*/
function leaderEmail(nombreLider,dataLideres)
{
  for (var i = 0 ; i < dataLideres.length ; i++)
  {
    if(nombreLider === dataLideres[i][0])
    {
      return dataLideres[i][1]
    }
  }
}

/*
FORMATEO DE FECHA
*/
function addLeadingZero(number) {
  return (number < 10 ? '0' : '') + number;
}
