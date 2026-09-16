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
  } else if (action === 'solicitar_edicion') {
    return cambiarEstadoEdicion(e, "Pendiente");
  } else if (action === 'aprobar_edicion') {
    return cambiarEstadoEdicion(e, "Abierto");
  } else if (action === 'verificar_edicion') {
    return verificarEdicion(e);
  }
  
  return ContentService.createTextOutput("Acción no válida").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  return doGet(e);
}

// 1. REGISTRAR ASISTENCIA (Con GPS)
function registrarAsistencia(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Registro");
    sheet.appendRow(["Fecha", "Hora", "Nombre", "Sucursal", "Tipo", "Ubicacion"]);
  }
  
  var fecha = e.parameter.fecha;
  var hora = e.parameter.hora;
  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var tipo = e.parameter.tipo; 
  var ubicacion = e.parameter.ubicacion || "Sin ubicación";
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("Ubicacion") === -1) {
    sheet.getRange(1, headers.length + 1).setValue("Ubicacion");
  }
  
  sheet.appendRow([fecha, hora, nombre, sucursal, tipo, ubicacion]);
  
  return ContentService.createTextOutput(JSON.stringify({"status": "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2. OBTENER DATOS
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
      obj[String(headers[j]).toLowerCase()] = row[j];
    }
    result.push(obj);
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 3. GUARDAR PERFIL (Y CERRAR ACCESO)
function guardarPerfil(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Perfiles");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Perfiles");
    sheet.appendRow(["Nombre", "Sucursal", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo", "EstadoEdicion"]);
  }
  
  var headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  if (headers.indexOf("EstadoEdicion") === -1) {
    sheet.getRange(1, headers.length + 1).setValue("EstadoEdicion");
    headers.push("EstadoEdicion");
  }

  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var horariosStr = e.parameter.horarios; 
  
  var horarios = {};
  try {
    horarios = JSON.parse(horariosStr);
  } catch (err) {}

  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
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
    horarios.domingo || "",
    "Cerrado" // Se bloquea al guardar
  ];
  
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 4. CAMBIAR ESTADO DE EDICIÓN
function cambiarEstadoEdicion(e, nuevoEstado) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Perfiles");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({"status": "error"})).setMimeType(ContentService.MimeType.JSON);
  
  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIndex = headers.indexOf("EstadoEdicion") + 1;
  
  if (colIndex === 0) {
      colIndex = headers.length + 1;
      sheet.getRange(1, colIndex).setValue("EstadoEdicion");
  }

  var success = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == nombre && data[i][1] == sucursal) {
      sheet.getRange(i + 1, colIndex).setValue(nuevoEstado);
      success = true;
      break;
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": success ? "ok" : "error"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 5. VERIFICAR EDICIÓN
function verificarEdicion(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Perfiles");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({"estado": "Cerrado"})).setMimeType(ContentService.MimeType.JSON);
  
  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIndex = headers.indexOf("EstadoEdicion");
  
  var estado = "Cerrado"; // por defecto
  
  if (colIndex > -1) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == nombre && data[i][1] == sucursal) {
        estado = data[i][colIndex] || "Cerrado";
        break;
      }
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({"estado": estado}))
    .setMimeType(ContentService.MimeType.JSON);
}

// 6. GUARDAR EXCEPCION
function guardarExcepcion(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Excepciones");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Excepciones");
    sheet.appendRow(["Nombre", "Sucursal", "Fecha", "Tipo"]);
  }
  var nombre = e.parameter.nombre;
  var sucursal = e.parameter.sucursal;
  var fecha = e.parameter.fecha;
  var tipo = e.parameter.tipo; 
  sheet.appendRow([nombre, sucursal, fecha, tipo]);
  return ContentService.createTextOutput(JSON.stringify({"status": "ok"})).setMimeType(ContentService.MimeType.JSON);
}

// 7. OBTENER PERFILES Y EXCEPCIONES
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
