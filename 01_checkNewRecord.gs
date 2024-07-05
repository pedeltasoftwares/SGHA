// Esta función identifica si hay una nueva línea que no esté en las hojas de Infraestructura, Edificaciones o Puentes
function checkNewRecord() {

  //Obtiene la hoja actual
  var book = SpreadsheetApp.getActiveSpreadsheet();

  //Hoja de entrada del formulario
  var sheet = book.getSheetByName("Respuestas de formulario 1");

  //Hoja con los correos de los líderes
  var sheetLideres = book.getSheetByName("Líderes");
  var dataLideres =  sheetLideres.getRange("A2:B").getValues();
        
  // Obtene todos los registros de la hoja del formulario
  var rango = sheet.getRange("B2:J");
  var registros = rango.getDisplayValues();

  // Usa el método filter para eliminar las posiciones que sean ""
  var registros = filterBlankSpacesSubarray(registros)

  //Itera sobre los registros
  for (var registro of registros) {
    
    // Usa el método filter para eliminar las posiciones que sean ""
    var registro = filterBlankSpacesArray(registro)
    
    //Formatea las fechas
    registro[1] = formatDateString(registro[1])

    //Aregla las posiciones del registro para que concida con el formato en la hoja de registro
    swap(registro,6,0)
    swap(registro,6,1)
    swap(registro,6,2)
    swap(registro,6,3)
    swap(registro,6,4)

    //Inserta "No" en el estado inicial de notificado
    registro.push("No")

    //Extrae la hoja del area
    var sheetArea = book.getSheetByName(registro[6]);

    //Registros escritos en la hoja del area
    var dataArea = sheetArea.getRange("A2:F").getDisplayValues();

    //Filtra los arrays que contengan registros
    var dataArea = filterBlankSpacesSubarray(dataArea)

    //Si no hay registro,agrega la fila
    if (dataArea.length === 0 )
    {
      //Elimina el área del registro
      registro.splice(6, 1)
      //Escribe en la hoja
      sheetArea.appendRow(registro)
      
    }
    else{

      //Elimina el área del registro
      registro.splice(6, 1)

      //Formatea la fecha a dd/mm/aaaa
      for (var i = 0; i < dataArea.length ; i++)
      {
        dataArea[i][2] = formatDateString(dataArea[i][2])
      }

      //Verifica si existe el registro en las hojas de cada dependencia
      var foundMatch = false;
      for (var data of dataArea)
      {
        if (data[0] === registro[0] && data[1] === registro[1] && data[2] === registro[2] && data[3] === registro[3] && data[4] === registro[4] && data[5] === registro[5])
        {
          foundMatch = true;
          break
        }
      }

      //Si no lo encuentra, lo escribe en la hoja correspondiente
      if (!foundMatch)
      {
        sheetArea.appendRow(registro)
      }
  
    }
      
  }
  // Envia la notificación al lider
  sendNotificationLeader(book,dataLideres)

}

/*
ELLIMINA LAS POSICIONES 0 QUE SEAN "" DE UN SUB ARRAY
*/
function filterBlankSpacesSubarray(array_temp)
{
  return array_temp.filter(function(subArray) {
    return subArray[0] !== "";
  });
}

/*
ELLIMINA LAS POSICIONES 0 QUE SEAN "" DE UN ARRAY
*/
function filterBlankSpacesArray(array_temp)
{
  return array_temp.filter(function(element) {
    return element !== "";
  });
}

  
/*
FORMATEA LA FECHA
*/
function formatDateString(dateString) {
  // Dividir la cadena de fecha en sus componentes
  var dateParts = dateString.split('/');
  
  // Asegurarse de que cada parte de la fecha tenga dos dígitos
  var day = ('0' + dateParts[0]).slice(-2);
  var month = ('0' + dateParts[1]).slice(-2);
  var year = dateParts[2];
  
  // Construir la cadena de fecha en el formato deseado
  var formattedDate = day + '/' + month + '/' + year;
  
  return formattedDate;
}

/*
CAMBIA DE POSICIÓN ELEMENTOS DE UN ARRAY
*/
function swap(array, index1, index2) {

  var temp = array[index1];
  array[index1] = array[index2];
  array[index2] = temp;

}
