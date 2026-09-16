function doGet(e) {
  var action = e.parameter.action;
  
  if (action === 'registrar') {
    return registrarAsistencia(e);
  } else if (action === 'obtener_datos') {
    return obtenerDatos();
  } else if (action === 'guardar_perfil') {
    return guardarPerfil(e);
  } else if (action === 'guardar_excepcion') {
    return guardarExcepcion(e);
  } else if (action === 'obtener_perfiles') {
    return obtenerPerfiles();
  }
  
  return ContentService.createTextOutput("Acción no válida").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  return doGet(e);
}

// 1. REGISTRAR ASISTENCIA (Reloj Checador)
function registrarAsistencia(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Registro");
    sheet.appendRow(["Fecha", "Hora", "Nombre", "Sucursal", "Tipo"]);
  }
  
  var fecha = e.parameter.fecha;
  var hora = e.parameter.hora;
  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var tipo = e.parameter.tipo; // Entrada o Salida
  
  sheet.appendRow([fecha, hora, nombre, sucursal, tipo]);
  
  return ContentService.createTextOutput(JSON.stringify({"status": "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2. OBTENER DATOS (Panel Directivo)
function obtenerDatos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    result.push(obj);
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 3. GUARDAR PERFIL Y HORARIOS
function guardarPerfil(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Perfiles");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Perfiles");
    sheet.appendRow(["Nombre", "Sucursal", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]);
  }
  
  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var horariosStr = e.parameter.horarios; // JSON string
  
  var horarios = {};
  try {
    horarios = JSON.parse(horariosStr);
  } catch (err) {
    // Error parseando json
  }

  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  // Buscar si ya existe el usuario
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == nombre && data[i][1] == sucursal) {
      rowIndex = i + 1;
      break;
    }
  }
  
  var rowData = [
    nombre, 
    sucursal,
    horarios.lunes || "",
    horarios.martes || "",
    horarios.miercoles || "",
    horarios.jueves || "",
    horarios.viernes || "",
    horarios.sabado || "",
    horarios.domingo || ""
  ];
  
  if (rowIndex > -1) {
    // Actualizar
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Insertar nuevo
    sheet.appendRow(rowData);
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 4. GUARDAR EXCEPCION (Falta/Permiso)
function guardarExcepcion(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Excepciones");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Excepciones");
    sheet.appendRow(["Nombre", "Sucursal", "Fecha", "Tipo"]);
  }
  
  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var fecha = e.parameter.fecha;
  var tipo = e.parameter.tipo; // "Falta", "Permiso", etc.
  
  sheet.appendRow([nombre, sucursal, fecha, tipo]);
  
  return ContentService.createTextOutput(JSON.stringify({"status": "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 5. OBTENER PERFILES Y EXCEPCIONES (Para el Panel)
function obtenerPerfiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetPerfiles = ss.getSheetByName("Perfiles");
  var perfiles = [];
  if (sheetPerfiles) {
    var dataP = sheetPerfiles.getDataRange().getValues();
    var headersP = dataP[0];
    for (var i = 1; i < dataP.length; i++) {
      var row = dataP[i];
      var obj = {};
      for (var j = 0; j < headersP.length; j++) {
        obj[headersP[j]] = row[j];
      }
      perfiles.push(obj);
    }
  }
  
  var sheetExcepciones = ss.getSheetByName("Excepciones");
  var excepciones = [];
  if (sheetExcepciones) {
    var dataE = sheetExcepciones.getDataRange().getValues();
    var headersE = dataE[0];
    for (var i = 1; i < dataE.length; i++) {
      var row = dataE[i];
      var obj = {};
      for (var j = 0; j < headersE.length; j++) {
        obj[headersE[j]] = row[j];
      }
      excepciones.push(obj);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "perfiles": perfiles,
    "excepciones": excepciones
  })).setMimeType(ContentService.MimeType.JSON);
}
