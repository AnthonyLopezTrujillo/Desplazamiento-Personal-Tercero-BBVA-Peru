const FILA_FORMULA_SOLICITUDES_IDX = 1;
const COLUMNA_FORMULA_SOLICITUDES_IDX = 16;

/**INI - ANT - 071124 */

const COLUMNA_FORMULA_TIEMPO_RESTANTE_IDX = 22;
const HORAS_RESTANTES_APROBACION = "16:00:00";

//INDICES SHEET FORM SOLICITUDES
const SHEET_RESPUESTA_SOLICITUDES_MARCA_TEMPORAL_IDX = 0;
const SHEET_RESPUESTA_SOLICITUDES_CORREO_REGISTRO_IDX = 1;
const SHEET_RESPUESTA_SOLICITUDES_NOMBRE_PROVEEDOR = 2;
const SHEET_RESPUESTA_SOLICITUDES_NOMBRE_EXTERNO_IDX = 3;
const SHEET_RESPUESTA_SOLICITUDES_DNI_EXTERNO_IDX = 4;
const SHEET_RESPUESTA_SOLICITUDES_AREA_EXTERNO_IDX = 5;
const SHEET_RESPUESTA_SOLICITUDES_PISO_HALL_EXTERNO_IDX = 6;
const SHEET_RESPUESTA_SOLICITUDES_NOMBRE_APELLIDO_COLABORADOR_IDX = 7;
const SHEET_RESPUESTA_SOLICITUDES_PREGISTRO_COLABORADOR_IDX = 8;
const SHEET_RESPUESTA_SOLICITUDES_FECHA_INGRESO_EXTERNO_IDX = 9;
const SHEET_RESPUESTA_SOLICITUDES_HORA_INGRESO_EXTERNO_IDX = 10;
const SHEET_RESPUESTA_SOLICITUDES_HORA_SALIDA_EXTERNO_IDX = 11;
const SHEET_RESPUESTA_SOLICITUDES_MARCA_INGRESO_REAL_IDX = 13;
const SHEET_RESPUESTA_SOLICITUDES_MARCA_SALIDA_REAL_IDX = 14;
const SHEET_RESPUESTA_SOLICITUDES_TIEMPO_DIARIO_IDX = 15;
const SHEET_RESPUESTA_SOLICITUDES_TIEMPO_SEMANAL_IDX = 16;
const SHEET_RESPUESTA_SOLICITUDES_ESTADO_IDX = 17;
const SHEET_RESPUESTA_SOLICITUDES_HORA_INGRESO_REAL_IDX = 18;
const SHEET_RESPUESTA_SOLICITUDES_HORA_SALIDA_REAL_IDX = 19;
const SHEET_RESPUESTA_SOLICITUDES_NUM_SEMANA_IDX = 20;
const SHEET_RESPUESTA_SOLICITUDES_TIEMPO_RESTANTEL_IDX = 22;

const CORREO_EQUIPO_RRLL = "";
const CORREO_EQUIPO_SF = "j";
const MAIL_SENDER_NAME = "";
const MAIL_SENDER_EMAIL = "";

const ESTADO_APROBADO = "APROBADO";
const ESTADO_DESAPROBADO = "DESAPROBADO";

const LISTA_CORREOS_N2_PROHIBIDOS = [""];

const LISTA_CORREOS_N1_MANAGER_PROHIBIDOS = [""];


//INDICES CAPACITY
const SHEET_CAPACITY_CORREO_COLABORADOR_IDX = 8;
const SHEET_CAPACITY_CORREO_MANAGER_IDX = 10;

function establecerFormulas(fila) {
  var formulaTiempoAcumulado = `=IF(P${fila} = ""; ""; SUMIFS(P:P; E:E; E${fila}; U:U; U${fila}))`;
  SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, COLUMNA_FORMULA_SOLICITUDES_IDX + 1).setFormula(formulaTiempoAcumulado);

  var formulaTiempoRestante = `=IF(U${fila}=""; ""; MAX(0; 16/24 - FILTER(Q:Q; E:E=E${fila}; U:U=U${fila}; ROW(Q:Q)=MIN(IF(E:E=E${fila}; IF(U:U=U${fila}; ROW(Q:Q); ))))))`;
  SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, COLUMNA_FORMULA_TIEMPO_RESTANTE_IDX + 1).setFormula(formulaTiempoRestante);
}

/*
function aplicarFormatos(fila) {
  var rangoHoraDiaria = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_TIEMPO_DIARIO_IDX + 1);
  var rangoHoraSemanal = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_TIEMPO_SEMANAL_IDX + 1);
  var rangoTiempoRestante = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_TIEMPO_RESTANTEL_IDX + 1);

  rangoHoraDiaria.setNumberFormat("HH:MM:SS");
  rangoHoraSemanal.setNumberFormat("[h]:mm:ss");
  rangoTiempoRestante.setNumberFormat("[h]:mm:ss");
}*/

function aplicarFormatos(fila) {
  var rangoHoraDiaria = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_TIEMPO_DIARIO_IDX + 1);
  var rangoHoraSemanal = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_TIEMPO_SEMANAL_IDX + 1);
  var rangoTiempoRestante = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_TIEMPO_RESTANTEL_IDX + 1);
  var rangoDni = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_DNI_EXTERNO_IDX + 1);
  var rangoNumeroSemana = SHEET_DESPLAZAMIENTO_TERCEROS.getRange(fila, SHEET_RESPUESTA_SOLICITUDES_NUM_SEMANA_IDX + 1);

  rangoNumeroSemana.setNumberFormat("0");
  rangoDni.setNumberFormat("@");
  rangoHoraDiaria.setNumberFormat("HH:MM:SS");
  rangoHoraSemanal.setNumberFormat("[h]:mm:ss");
  rangoTiempoRestante.setNumberFormat("[h]:mm:ss");
}

function obtenerNumeroSemana(fechaTexto) {
  var partesFecha = fechaTexto.split('/');
  var dia = parseInt(partesFecha[0], 10);
  var mes = parseInt(partesFecha[1], 10) - 1;
  var anio = parseInt(partesFecha[2], 10);
  var fecha = new Date(anio, mes, dia);

  var primerDiaDelAno = new Date(anio, 0, 1);
  var diaSemanaPrimerDia = primerDiaDelAno.getDay() || 7;
  primerDiaDelAno.setDate(primerDiaDelAno.getDate() + (1 - diaSemanaPrimerDia));

  var diferenciaDias = Math.floor((fecha - primerDiaDelAno) / (24 * 60 * 60 * 1000));
  var numeroSemana = Math.ceil((diferenciaDias + 1) / 7);
  return numeroSemana.toString();
}


function formatearFechaATexto(fecha) {
  var dia = fecha.getDate().toString().padStart(2, '0');
  var mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
  var anio = fecha.getFullYear();
  return `${dia}/${mes}/${anio}`;
}

function verificarAprobacion(duracionTexto) {
  var partes = duracionTexto.split(':');
  var horas = parseInt(partes[0], 10);
  var minutos = parseInt(partes[1], 10);
  var segundos = parseInt(partes[2], 10);

  if (horas >= 16) {
    return false;
  } else if (horas === 16) {
    if (minutos > 0 || segundos > 0) {
      return false;
    }
  }

  return true;
}
/*
function sumarHoras(duraciones) {
  var totalSegundos = 0;

  duraciones.forEach(function(duracionTexto) {
    var partes = duracionTexto.split(':');
    var horas = parseInt(partes[0], 10);
    var minutos = parseInt(partes[1], 10);
    var segundos = parseInt(partes[2], 10);
    
    totalSegundos += (horas * 3600) + (minutos * 60) + segundos;
  });

  var horasTotales = Math.floor(totalSegundos / 3600);
  var minutosTotales = Math.floor((totalSegundos % 3600) / 60);
  var segundosTotales = totalSegundos % 60;

  return `${horasTotales.toString().padStart(2, '0')}:${minutosTotales.toString().padStart(2, '0')}:${segundosTotales.toString().padStart(2, '0')}`;
}
*/

function sumarHoras(duraciones) {
  var totalSegundos = 0;

  duraciones.forEach(function (duracionTexto) {
    // Verificar que la duración esté en formato hh:mm:ss
    if (duracionTexto && duracionTexto.includes(":")) {
      var partes = duracionTexto.split(':');
      if (partes.length === 3) {
        var horas = parseInt(partes[0], 10) || 0;
        var minutos = parseInt(partes[1], 10) || 0;
        var segundos = parseInt(partes[2], 10) || 0;

        totalSegundos += (horas * 3600) + (minutos * 60) + segundos;
      } else {
        console.warn(`Formato inválido de duración: ${duracionTexto}`);
      }
    } else {
      console.warn(`Duración vacía o en formato incorrecto: ${duracionTexto}`);
    }
  });

  var horasTotales = Math.floor(totalSegundos / 3600);
  var minutosTotales = Math.floor((totalSegundos % 3600) / 60);
  var segundosTotales = totalSegundos % 60;

  return `${horasTotales.toString().padStart(2, '0')}:${minutosTotales.toString().padStart(2, '0')}:${segundosTotales.toString().padStart(2, '0')}`;
}

function restarHoras(duracionTexto) {
  // Convertimos 16:00:00 a segundos
  const horasLimite = 16 * 3600; // 16 horas en segundos

  // Dividimos duracionTexto en horas, minutos y segundos
  const partes = duracionTexto.split(':');
  const horas = parseInt(partes[0], 10) || 0;
  const minutos = parseInt(partes[1], 10) || 0;
  const segundos = parseInt(partes[2], 10) || 0;

  // Convertimos duracionTexto a segundos
  const duracionSegundos = (horas * 3600) + (minutos * 60) + segundos;

  // Calculamos la diferencia en segundos, asegurándonos de que no sea negativa
  const diferenciaSegundos = Math.max(0, horasLimite - duracionSegundos);

  // Convertimos la diferencia a horas, minutos y segundos
  const horasResultado = Math.floor(diferenciaSegundos / 3600);
  const minutosResultado = Math.floor((diferenciaSegundos % 3600) / 60);
  const segundosResultado = diferenciaSegundos % 60;

  // Retornamos el resultado en formato hh:mm:ss
  return `${horasResultado.toString().padStart(2, '0')}:${minutosResultado.toString().padStart(2, '0')}:${segundosResultado.toString().padStart(2, '0')}`;
}



function obtenerFilaConFechaMasAntigua(dataExterno) {
  console.log('Datos de entrada:', dataExterno);

  const indiceFecha = SHEET_RESPUESTA_SOLICITUDES_HORA_SALIDA_REAL_IDX;

  if (dataExterno.length === 1) {
    return dataExterno[0];
  }

  let filaMasAntigua = dataExterno[0];
  let fechaMasAntigua = parseFecha(dataExterno[0][indiceFecha]);

  for (let i = 1; i < dataExterno.length; i++) {
    let fechaActual = parseFecha(dataExterno[i][indiceFecha]);

    if (fechaActual > fechaMasAntigua) {
      fechaMasAntigua = fechaActual;
      filaMasAntigua = dataExterno[i];
    }
  }

  return filaMasAntigua;
}

function parseFecha(fechaTexto) {
  const [fecha, hora] = fechaTexto.split(" ");
  const [dia, mes, anio] = fecha.split("/").map(num => parseInt(num, 10));
  const [horas, minutos, segundos] = hora.split(":").map(num => parseInt(num, 10));

  const parsedDate = new Date(anio, mes - 1, dia, horas, minutos, segundos);

  return parsedDate;
}

function calcularHoraAcumuladas(dataExterno) {
  var duraciones = dataExterno.map(function (item) {
    return item[SHEET_RESPUESTA_SOLICITUDES_TIEMPO_DIARIO_IDX];
  });
  var horasAcumuladas = sumarHoras(duraciones);
  return horasAcumuladas;
}

function calcularHoraRestantes(horasAcumuladas) {
  var resta = restarHoras(horasAcumuladas);
  return resta;

}

function formatearTiempo(tiempo) {
  const [horas, minutos, segundos] = tiempo.split(':');
  let resultado = "";

  resultado += `${horas} hora${horas !== '1' ? 's' : ''}`;
  resultado += `, ${minutos} minuto${minutos !== '1' ? 's' : ''}`;
  resultado += ` y ${segundos} segundo${segundos !== '1' ? 's' : ''}`;
  return resultado;
}

function hallarCorreoManager(dataCapacity, correoEnvio) {
  var correoColaborador = correoEnvio.toUpperCase().trim();
  console.log("correoColaborador:: " + correoColaborador);

  for (var i = 0; i < dataCapacity.length; i++) {
    var itemCapacity = dataCapacity[i];
    var correoColaboradorCapacity = itemCapacity[SHEET_CAPACITY_CORREO_COLABORADOR_IDX];
    var correoManager = itemCapacity[SHEET_CAPACITY_CORREO_MANAGER_IDX];

    if (correoColaboradorCapacity != "") {
      correoColaboradorCapacity = correoColaboradorCapacity.toUpperCase().trim();
    }

    if (correoColaboradorCapacity == correoColaborador) {
      if (correoManager != "") {
        correoManager = correoManager.toUpperCase().trim();
      }
      return correoManager;
    }
  }

  return null;
}

/**FIN - ANT - 071124 */
