function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gestión Edificio MECM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Obtener datos con limpieza de números
function obtenerDatosEdificio() {
  try {
    // Usar la hoja actual donde está desplegado el script
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('BD edificio MECM');
    
    if (!sheet) {
      // Si no encuentra la hoja con ese nombre, usa la primera hoja
      sheet = ss.getSheets()[0];
    }
    
    var data = sheet.getDataRange().getValues();
    var objetoEdificio = {};

    var limpiarNum = function(valor) {
      if (!valor && valor !== 0) return 0;
      if (typeof valor === 'number') return valor;
      var procesado = String(valor).replace(/[^0-9.-]/g, '');
      var numero = parseFloat(procesado);
      return isNaN(numero) ? 0 : numero;
    };

    for (var i = 1; i < data.length; i++) {
      var fila = data[i];
      var idDepto = String(fila[1]).trim();
      if (idDepto && idDepto !== "") {
        objetoEdificio[idDepto] = {
  inquilino: String(fila[2] || ""),
  carnet: String(fila[3] || ""),
  fono: String(fila[4] || ""),

  alquiler: limpiarNum(fila[5]),
  mantenimiento: limpiarNum(fila[6]),
  limpieza: limpiarNum(fila[7]),
  seguridad: limpiarNum(fila[8]),
  agua: limpiarNum(fila[9]),
  electricidad: limpiarNum(fila[10]),

  totalMora: limpiarNum(fila[11]),

  baucher: String(fila[12] || ""),
  montoRecibido: limpiarNum(fila[13]),
  diferencia: limpiarNum(fila[14]),

  estadoSheet: String(fila[15] || "Con deuda").trim()
};
      }
    }
    
    Logger.log("Datos cargados: " + Object.keys(objetoEdificio).length + " departamentos");
    return objetoEdificio;
    
  } catch (error) {
    Logger.log("Error en obtenerDatosEdificio: " + error.toString());
    return {};
  }
}

// 2. Función para guardar archivo en Drive y actualizar Excel
function registrarPagoConArchivo(e) {
  try {
    if (!e || !e.datosBase64 || !e.idDepto) {
      throw new Error("Datos incompletos del comprobante");
    }

    // Crear carpeta "Vouchers Pago" si no existe
    var carpetas = DriveApp.getFoldersByName("Vouchers Pago Edificio");
    var carpeta;
    if (carpetas.hasNext()) {
      carpeta = carpetas.next();
    } else {
      carpeta = DriveApp.createFolder("Vouchers Pago Edificio");
    }

    // Convertir y guardar el archivo
    var bytes = Utilities.base64Decode(e.datosBase64);
    var blob = Utilities.newBlob(bytes, e.tipoMime, e.nombreArchivo);
    var archivoGuardado = carpeta.createFile(blob);
    archivoGuardado.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var urlDeLaFoto = archivoGuardado.getUrl();

    // Conectar a la hoja de cálculo
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('BD edificio MECM');
    
    if (!hoja) {
      hoja = ss.getSheets()[0];
    }
    
    var datos = hoja.getDataRange().getValues();
    var filaEncontrada = -1;

    // Buscar el departamento (columna B)
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][1].toString().trim() === e.idDepto.toString().trim()) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) {
      throw new Error("No se encontró el Departamento " + e.idDepto);
    }

    // Actualizar: Columna M (13) = URL del comprobante
    hoja.getRange(filaEncontrada, 13).setValue(urlDeLaFoto);

    return {
      exito: true,
      mensaje: "Pago registrado exitosamente",
      urlFoto: urlDeLaFoto
    };

  } catch (error) {
    throw new Error("Error: " + error.message);
  }
}
