/*
Envía el reporte semanal a los líderes de cada área. Se enviará el viernes a las 2 PM
*/
function sendWeeklyReportLeader() {

  //Obtiene la hoja actual
  var book = SpreadsheetApp.getActiveSpreadsheet();

  //Obtiene las dependencias y los correos a cuales enviar el reporte
  var sheet = book.getSheetByName("Informe semanal");
  var areas =  sheet.getRange("A1:A3").getValues();
  var correos =  sheet.getRange("B1:B3").getValues();
  areas = [areas[0][0], areas[1][0],areas[2][0]]
  correos = [correos[0][0], correos[1][0],correos[2][0]]

  //Fecha del lunes y del viernes de la semana actual
  var friday = new Date();
  var dayOfWeek = friday.getDay(); 

  var monday = new Date(friday);
  monday.setDate(friday.getDate() - (dayOfWeek - 1));

  // Itera sobre las horas para
  for (var area of areas)
  {
    //Obtiene la hoja
    sheet = book.getSheetByName(area);

    //Obtiene todos los registros de la hoja del área
    var registros = sheet.getRange("A2:I").getDisplayValues();

    //Filtra los arrays que contengan registros
    var registros = filterBlankSpacesSubarray(registros)

    //Formatea la fecha a dd/mm/aaaa
    for (var i = 0; i < registros.length ; i++)
    {
      registros[i][2] = formatDateString(registros[i][2])
    }

    //Inicializa el array donde almacena los registros semanales para el reporte
    var report_records = []

    //Itera por los registros para encontrar las solicitudes de esta semana
    for (var i = 0 ; i < registros.length;i++)
    {
      //Convierte la fecha a date
      var date = stringToDate(registros[i][2])

      if (date >= monday && date <= friday) {
        report_records.push([registros[i][0], registros[i][1], registros[i][2],registros[i][3],registros[i][4],registros[i][5],registros[i][8]])
      }

    }

    //Envía el reporte
    if (report_records.length >= 1){

      //Correo del lider
      if(area === "Infraestructura")
      {
        var correo_lider =  correos[0]
      }
      else if ( area === "Puentes" )
      {
        var correo_lider =  correos[1]
      }
      else
      {
        var correo_lider =  correos[2]
      }
      
      //Enviar reporte
      sendReport(correo_lider, report_records, friday, monday)
    }

  }

}

/*
Convierte los strings a fechas
*/
function stringToDate(fechaStr) {

  var partes = fechaStr.split('/');
  var dia = parseInt(partes[0], 10);
  var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript son 0-11
  var anio = parseInt(partes[2], 10);
  return new Date(anio, mes, dia);
}

/*
Envía el reporte
*/
function sendReport(correo_lider, report_records, friday, monday) {

  //Formatea las fechas a string
  monday = dateToString(monday)
  friday = dateToString(friday)

  //Agrega el encabezado
  report_records.unshift(["Nombre", "Correo", "Fecha hora adicional", "Proyecto", "Descripción labor", "Líder", "Aprobado"]);

  // Construir la tabla en HTML
  var htmlTable = '<table border="1" style="border-collapse: collapse; width: 60%;">';
  for (var i = 0; i < report_records.length; i++) {
    htmlTable += '<tr>';
    for (var j = 0; j < report_records[i].length; j++) {
      if (i === 0) {
        htmlTable += '<th style="padding: 4px; text-align: left;">' + report_records[i][j] + '</th>';
      } else {
        htmlTable += '<td style="padding: 4px; text-align: left;">' + report_records[i][j] + '</td>';
      }
    }
    htmlTable += '</tr>';
  }
  htmlTable += '</table>';
  
  // Contenido del correo electrónico
  var emailContent = 'Estimado/a,<br><br>Adjunto encontrará el reporte de horas adicionales correspondientes al período del '+ monday + ' al ' + friday + ', realizado por su equipo de trabajo.<br><br>' + htmlTable + '<br><br>Atentamente,<br>Sistema de Gestión de Horas Adicionales';

  // Enviar el correo electrónico
  GmailApp.sendEmail(correo_lider, 'Reporte Semanal Horas Adicionales - ' + monday + ' al ' + friday, '', {
    htmlBody: emailContent, from: "horas-adicionales-notificaciones@pedelta.com.co"
  });
}

/*
Date to string
*/
function dateToString(fecha) {

  // Obtener los componentes de la fecha
  var dia = fecha.getDate();
  var mes = fecha.getMonth() + 1; // Los meses en JavaScript son 0-11
  var anio = fecha.getFullYear();

  // Asegurar que el día y el mes tengan dos dígitos
  if (dia < 10) {
    dia = '0' + dia;
  }
  if (mes < 10) {
    mes = '0' + mes;
  }

  // Formatear la fecha como cadena en el formato "dd/MM/yyyy"
  var fechaCadena = dia + '/' + mes + '/' + anio;

  return fechaCadena
}