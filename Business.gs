function triggerFormRespuestas(e) {
  var ultimaFila = SHEET_DESPLAZAMIENTO_TERCEROS.getLastRow();
  establecerFormulas(ultimaFila);
  aplicarFormatos(ultimaFila);

  var valoresUltimaFila = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(ultimaFila, 1, 1, SHEET_DESPLAZAMIENTO_TERCEROS.getLastColumn()).getValues()[0];
  var correoEnvio = valoresUltimaFila[SHEET_RESPUESTA_SOLICITUDES_CORREO_REGISTRO_IDX];
  var numDocumento = valoresUltimaFila[SHEET_RESPUESTA_SOLICITUDES_DNI_EXTERNO_IDX];
  var fechaFormIngreso = valoresUltimaFila[SHEET_RESPUESTA_SOLICITUDES_FECHA_INGRESO_EXTERNO_IDX];
  console.log("correoEnvio:: " + correoEnvio + " ::numDocumento:: " + numDocumento+ " ::fechaFormIngreso:: " + fechaFormIngreso);
  var fechaTrans = formatearFechaATexto(fechaFormIngreso);
  var numSemana = obtenerNumeroSemana(fechaTrans);

  var dataExterno = obtenerDataExterno(numDocumento, numSemana);
  var estadoSolicitud = verificarSolicitud(dataExterno);

  console.log("fechaATexto:: " + fechaTrans + " numSemana:: " + numSemana+"dataExterno:: " + dataExterno.length + " estadoSolicitud:: " + estadoSolicitud);

  armarCorreoEstadoSolicitud(dataExterno,estadoSolicitud,correoEnvio);
  registrarEstadoEnSheet(ultimaFila,estadoSolicitud);
}

