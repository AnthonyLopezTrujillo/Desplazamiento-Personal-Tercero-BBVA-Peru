function armarCorreoEstadoSolicitud(dataExterno, aprobacionSolicitud, correoEnvio) {
    if (dataExterno && dataExterno.length >= 1) {
        var horasAcumuladas = calcularHoraAcumuladas(dataExterno);
        var horasRestantes = calcularHoraRestantes(horasAcumuladas);
        console.log("horasAcumuladas:: " + horasAcumuladas + " horasRestantes:: " + horasRestantes);
        if (aprobacionSolicitud == true) {
            enviarCorreoAprobacion(horasRestantes, correoEnvio);
        } else {
            var dataCapacity = obtenerDataCapacity();
            console.log("dataCapacity:: " + dataCapacity.length);
            var correoManager = hallarCorreoManager(dataCapacity, correoEnvio);
            console.log("correoManager:: " + correoManager);
            enviarCorreoRechazo(horasAcumuladas, correoEnvio, correoManager);
        }

    } else {
        var horasRestantes = HORAS_RESTANTES_APROBACION;
        enviarCorreoAprobacion(horasRestantes, correoEnvio);
    }
}

function enviarCorreoAprobacion(horasRestantes, correoEnvio) {
    var asunto = "Solicitud de Desplazamiento Temporal Terceros - APROBADA";
    var emailPlantillaRespuesta = HtmlService.createTemplateFromFile('Correos/SolicitudAprobacion');
    emailPlantillaRespuesta.horasRestantes = formatearTiempo(horasRestantes);
    var htmlMessageRespuesta = emailPlantillaRespuesta.evaluate().getContent();
    enviarCorreoHtml(correoEnvio, asunto, htmlMessageRespuesta);
}

function enviarCorreoRechazo(horasAcumuladas, correoEnvio, correoManager) {
    var asunto = "Solicitud de Desplazamiento Temporal Terceros - RECHAZADA";
    var emailPlantillaRespuesta = HtmlService.createTemplateFromFile('Correos/SolicitudRechazo');
    emailPlantillaRespuesta.horasAcumuladas = formatearTiempo(horasAcumuladas);
    var htmlMessageRespuesta = emailPlantillaRespuesta.evaluate().getContent();

    var emailDestino = correoEnvio.toUpperCase().trim();
    if (LISTA_CORREOS_N2_PROHIBIDOS.includes(emailDestino)) {
        emailDestino = ""; 
    }
    if (LISTA_CORREOS_N1_MANAGER_PROHIBIDOS.includes(emailDestino)) {
        emailDestino = ""; 
    }

  if (correoManager != null) {
    var correoCopia = correoManager.toUpperCase().trim();
    if (LISTA_CORREOS_N2_PROHIBIDOS.includes(correoCopia)) {
      correoCopia = "";
    }
    if (LISTA_CORREOS_N1_MANAGER_PROHIBIDOS.includes(correoCopia)) {
      correoCopia = "";
    }
  }else if(correoManager == null){
    var correoCopia = "";

  }

    enviarCorreoRechazoHtmlConCC(emailDestino, asunto, htmlMessageRespuesta, correoCopia);
}

function enviarCorreoHtml(emailDestino, asunto, mensaje) {
    var options = {
        htmlBody: mensaje,
        name: MAIL_SENDER_NAME,
        from: MAIL_SENDER_EMAIL
    };
    GmailApp.sendEmail(emailDestino, asunto, mensaje, options);
}

function enviarCorreoRechazoHtmlConCC(emailDestino, asunto, mensaje, copia) {
    var options = {
        htmlBody: mensaje,
        cc: CORREO_EQUIPO_RRLL + ", " + CORREO_EQUIPO_SF + ", " + copia,
        name: MAIL_SENDER_NAME,
        from: MAIL_SENDER_EMAIL
    };
    GmailApp.sendEmail(emailDestino, asunto, mensaje, options);


}

function enviarCorreoHtmlConCC(emailDestino, asunto, mensaje) {
    var options = {
        htmlBody: mensaje,
        cc: CORREO_EQUIPO_RRLL + ", " + CORREO_EQUIPO_SF,
        name: MAIL_SENDER_NAME,
        from: MAIL_SENDER_EMAIL
    };
    GmailApp.sendEmail(emailDestino, asunto, mensaje, options);
}
