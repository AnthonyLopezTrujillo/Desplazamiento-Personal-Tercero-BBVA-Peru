/** INI - ANT - 081124 */

// Data Colaboradores Workday (Activos y futuros)
const SS_DESPLAZAMIENTO_TERCEROS_ID = "";
const SS_DESPLAZAMIENTO_TERCEROS_SS = SpreadsheetApp.openById(SS_DESPLAZAMIENTO_TERCEROS_ID);
const SHEET_DESPLAZAMIENTO_NOMBRE = "Respuestas de formulario 1";
const SHEET_DESPLAZAMIENTO_TERCEROS = SS_DESPLAZAMIENTO_TERCEROS_SS.getSheetByName(SHEET_DESPLAZAMIENTO_NOMBRE);

//Data Capacity
const SS_CAPACITY_ID = "";
const SHEET_CAPACITY_NOMBRE = "BASE";

function obtenerDataExterno(numeroIdentidad, numSemana) {
  var query = "SELECT * WHERE E = '" + numeroIdentidad + "' AND U = " + numSemana + "";
  console.log("queryFormula::: " + query);
  var resultado = executeQuery(SS_DESPLAZAMIENTO_TERCEROS_ID, SHEET_DESPLAZAMIENTO_NOMBRE, query, "A1:W", false);
  return resultado;
}

function obtenerDataCapacity() {
  var query = "SELECT * WHERE G <> '' ";
  console.log("queryFormula Capacity::: " + query);
  var resultado = executeQuery(SS_CAPACITY_ID, SHEET_CAPACITY_NOMBRE, query, "A1:K", false);
  return resultado;
}

function verificarSolicitud(dataExterno) {
  var verificacion = false;

  if (dataExterno && dataExterno.length > 0) {
    console.log("dataExterno tiene más de 0 elementos y es distinto de null");

    if (dataExterno.length === 1) {
      var filaData = dataExterno[0];
      console.log(filaData); 
      console.log("dataExterno tiene exactamente 1 elemento");
      var horaAcumulada = filaData[SHEET_RESPUESTA_SOLICITUDES_TIEMPO_SEMANAL_IDX];
      console.log("horaAcumulada:: " + horaAcumulada);
      verificacion = verificarAprobacion(horaAcumulada);
      return verificacion;

    } else if (dataExterno.length > 1) {
      console.log("dataExterno tiene más de 1 elemento");
      var duraciones = dataExterno.map(function (item) {
        return item[SHEET_RESPUESTA_SOLICITUDES_TIEMPO_DIARIO_IDX];
      });
      console.log("Listaduraciones:: " + duraciones.length);
      var duracionAcumuladaTexto = sumarHoras(duraciones);
      console.log("duracionAcumuladaTexto:: " + duracionAcumuladaTexto);
      verificacion = verificarAprobacion(duracionAcumuladaTexto);
      console.log("verificarAprobacion:: " + verificacion);
      return verificacion;

    }
  } else {
    console.log("no hay registros");
    verificacion = true;
    return verificacion
  }

}

function registrarEstadoEnSheet(ultimaFila, estadoSolicitud) {
  if (estadoSolicitud == true) {
    SHEET_DESPLAZAMIENTO_TERCEROS.getRange(ultimaFila, SHEET_RESPUESTA_SOLICITUDES_ESTADO_IDX + 1, 1, 1).setValue(ESTADO_APROBADO);
  } else {
    SHEET_DESPLAZAMIENTO_TERCEROS.getRange(ultimaFila, SHEET_RESPUESTA_SOLICITUDES_ESTADO_IDX + 1, 1, 1).setValue(ESTADO_DESAPROBADO);
  }
}


/****************************/
/***** QUERYS GENERALES *****/
/****************************/
// Query generico para busquedas en archivos
function executeQuery(spreadSheetId, sheetName, queryFormula, range, showHeader) {
  var lastRow = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName).getLastRow();
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + spreadSheetId
    + '/gviz/tq?tqx=out:json&headers=1&sheet=' + sheetName
    + '&range=' + range + lastRow
    + '&tq=' + encodeURIComponent(queryFormula);
  //console.log("qvizURL=" + qvizURL);
  var ret = UrlFetchApp.fetch(qvizURL, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getContentText();
  var resp = JSON.parse(ret.replace("/*O_o*/", "").replace("google.visualization.Query.setResponse(", "").slice(0, -2));
  var data = resp.table.rows.map(row => {
    return row.c.map(cols => {
      return cols === null ? '' : cols.f !== undefined ? cols.f : (cols.v === null ? '' : cols.v);
    });
  });
  if (showHeader) {
    var header = resp.table.cols.map(col => {
      return col.label;
    });
    return [header].concat(data);
  } else {
    return data;
  }
}

/** FIN - ANT - 081124 */
