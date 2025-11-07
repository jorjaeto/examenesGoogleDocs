/**
 * =================================================================================
 * NOTA SOBRE PERMISOS
 * =================================================================================
 * Este script utiliza un archivo de manifiesto 'appsscript.json' para declarar
 * explícitamente los permisos (OAuth Scopes) necesarios para funcionar. Esto es
 * crucial para que la aplicación web pueda acceder a Google Sheets, Drive y Mail
 * en tu nombre. Al implementar la aplicación, se te pedirá que autorices estos
 * alcances.
 * =================================================================================
 */

// =================================================================================
// CONFIGURACIÓN INICIAL - ¡DEBES MODIFICAR ESTOS VALORES!
// =================================================================================
const ID_HOJA_REGISTRO = '1f2x20PuULCQDhm_568gdyGzdmlTRlXSqGo1-4Jp14oU'; // Reemplaza con el ID de tu Google Sheet para registros.
const ID_CARPETA_EXAMENES = '17RX0-kVYTyHDcVaE1CnL5bXmiPqi7bsh'; // Reemplaza con el ID de la carpeta de Drive donde se guardarán los exámenes.
const ID_PLANTILLA_DOC = '1qLuR2Asek8B1uGrWSWtkHJic9ua1Xqdi0PaY677fg74'; // <-- ¡IMPORTANTE! Reemplaza con el ID de tu plantilla de Google Docs

// Nombres de las hojas dentro del Google Sheet de registro.
const NOMBRE_HOJA_REGISTRO = 'Registros';
const NOMBRE_HOJA_CONFIG = 'Configuracion';
// =================================================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index').setTitle("Portal de Examen");
}

function doPost(e) {
  return procesarFormulario(e.parameter);
}

/**
 * Procesa la lógica principal del formulario. Ahora lanza una excepción en caso de error
 * para ser capturada por withFailureHandler en el lado del cliente.
 * @param {Object} form Objeto con los datos del formulario enviado.
 * @return {Object} Un objeto con los datos de éxito.
 */
function procesarFormulario(form) {
  const lock = LockService.getScriptLock();
  try {
    // Aumentado el tiempo de espera a 60 segundos. Si falla, lanzará una excepción.
    lock.waitLock(60000);

    const hojaRegistro = SpreadsheetApp.openById(ID_HOJA_REGISTRO).getSheetByName(NOMBRE_HOJA_REGISTRO);
    if (!hojaRegistro) {
      throw new Error(`La hoja de registro con el nombre "${NOMBRE_HOJA_REGISTRO}" no existe.`);
    }

    const tipoEnvio = form.tipoEnvio;
    const codigo = form.codigo1 || form.codigoContinuar || form.codigoEntregar;

    if (!estaHabilitado(form.turno)) {
      throw new Error(`Los envíos están actualmente deshabilitados para el ${form.turno}.`);
    }

    const datosUsuario = obtenerDatosExamen(codigo, hojaRegistro);
    console.log("Datos de usuario " + datosUsuario)
    if (datosUsuario && datosUsuario.entregado) {
      throw new Error("Este ejercicio ya ha sido entregado previamente y no puede ser modificado.");
    }

    let resultado;
    switch (tipoEnvio) {
      case "Iniciar":
        resultado = iniciarExamen(form, hojaRegistro, datosUsuario);
        break;
      case "Continuar":
        resultado = continuarExamen(form, hojaRegistro, datosUsuario);
        break;
      case "Entregar":
        resultado = entregarExamen(form, hojaRegistro, datosUsuario);
        break;
      default:
        throw new Error("Tipo de envío no válido.");
    }
    return resultado;

  } catch (e) {
    // **MODIFICADO**: Captura CUALQUIER error (timeout, Drive, etc.), lo registra y lo lanza de nuevo.
    // Esto es crucial para que withFailureHandler en el cliente lo reciba.
    Logger.log(`Error en procesarFormulario: ${e.toString()} - Stack: ${e.stack}`);
    throw e;
  } finally {
    // Este bloque SIEMPRE se ejecuta, asegurando que el bloqueo se libere.
    lock.releaseLock();
  }
}

function iniciarExamen(form, hojaRegistro, datosUsuario) {
  if (datosUsuario) {
    throw new Error("Este código de usuario ya ha iniciado la prueba. Por favor, utilice la opción 'Continuar' o 'Entregar'.");
  }
  const { codigo1: codigo, contrasena1: contrasena, turno, aula } = form;
  const fechaHora = new Date();
  const carpetaPrincipal = DriveApp.getFolderById(ID_CARPETA_EXAMENES);
  const plantillaDoc = DriveApp.getFileById(ID_PLANTILLA_DOC);
  const carpetaOpositor = carpetaPrincipal.createFolder(codigo);
  const nombreArchivoDoc = `${codigo}-${turno}-${aula}-Ejercicio`.trim();
  const nuevoDocumento = plantillaDoc.makeCopy(nombreArchivoDoc, carpetaOpositor);
  nuevoDocumento.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  const urlDoc = nuevoDocumento.getUrl();
  const idDoc = nuevoDocumento.getId();
  
  hojaRegistro.appendRow([fechaHora, "'" + codigo, "'" + contrasena, turno, aula, "Iniciar", idDoc, urlDoc, "", false]);
  return {
    mensaje: `Se ha generado el ejercicio para el aspirante con código ${codigo}. A continuación tienes el enlace para acceder:`,
    enlaces: [
        { nombre: "Ejercicio (Google Docs)", url: urlDoc },
    ]
  };
}

function continuarExamen(form, hojaRegistro, datosUsuario) {
  if (!datosUsuario) {
    throw new Error("No existe registro de inicio de examen para este código. Por favor, utilice la opción 'Iniciar'.");
  }
  if (String(datosUsuario.contrasena) !== String(form.contrasenaContinuar)) {
    throw new Error("La contraseña introducida no es correcta. Verifíquela y vuelva a intentarlo.");
  }
  hojaRegistro.appendRow([new Date(), "'" + form.codigoContinuar, "'" + form.contrasenaContinuar, datosUsuario.turno, datosUsuario.aula, "Continuar", datosUsuario.idArchivoDoc, datosUsuario.urlArchivoDoc, "", false]);
  return {
    mensaje: "Puede continuar su ejercicio en los siguientes enlaces:",
    enlaces: [
        { nombre: "Ejercicio (Google Docs)", url: datosUsuario.urlArchivoDoc },
    ]
  };
}

function entregarExamen(form, hojaRegistro, datosUsuario) {
  if (!datosUsuario) {
    throw new Error("No existe registro de inicio de examen para este código. No se puede entregar.");
  }
  if (String(datosUsuario.contrasena) !== String(form.contrasenaEntregar)) {
    throw new Error("La contraseña introducida no es correcta.");
  }
  const archivoDoc = DriveApp.getFileById(datosUsuario.idArchivoDoc);
  const carpetaOpositor = archivoDoc.getParents().next();
  try {
    archivoDoc.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    archivoDoc.getEditors().forEach(editor => archivoDoc.removeEditor(editor));
    archivoDoc.getViewers().forEach(viewer => archivoDoc.removeViewer(viewer));
  } catch (e) {
    Logger.log(`No se pudieron remover todos los permisos del archivo Doc ${datosUsuario.idArchivoDoc}. Error: ${e}`);
  }
  exportarYGuardarEntregables(archivoDoc, carpetaOpositor, form.codigoEntregar, form.correoEntregar);
  hojaRegistro.getRange(datosUsuario.fila, 10).setValue(true);
  hojaRegistro.appendRow([new Date(), "'" + form.codigoEntregar, "'" + form.contrasenaEntregar, datosUsuario.turno, datosUsuario.aula, "Entregar", datosUsuario.idArchivoDoc, datosUsuario.urlArchivoDoc, form.correoEntregar, true]);
  return {
    mensaje: `El ejercicio con código ${form.codigoEntregar} ha sido entregado correctamente. Si ha proporcionado un correo electrónico, recibirá una copia. Ya puede cerrar esta ventana.`,
    enlaces: []
  };
}

function obtenerDatosExamen(codigo, hojaRegistro) {
  const datos = hojaRegistro.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) {
    if (String(datos[i][1]) === String(codigo) && String(datos[i][5]) === String("Iniciar")) {
      return {
        fila: i + 1,
        contrasena: datos[i][2],
        turno: datos[i][3],
        aula: datos[i][4],
        idArchivoDoc: datos[i][6],
        urlArchivoDoc: datos[i][7],
        entregado: datos[i][9] === true
      };
    }
  }
  return null;
}

function estaHabilitado(turno) {
  const hojaConfig = SpreadsheetApp.openById(ID_HOJA_REGISTRO).getSheetByName(NOMBRE_HOJA_CONFIG);
  if (!hojaConfig) {
     Logger.log(`ADVERTENCIA: La hoja de configuración "${NOMBRE_HOJA_CONFIG}" no existe. Se permiten todos los envíos por defecto.`);
     return true;
  }
  const datos = hojaConfig.getRange("A1:B2").getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === turno) {
      return datos[i][1] === true;
    }
  }
  return false;
}

function exportarYGuardarEntregables(archivo, carpetaDestino, codigo, correo) {
  try {
    const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${archivo.getId()}&exportFormat=pdf`;
    const opciones = {
      muteHttpExceptions: true,
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
    };
    const respuesta = UrlFetchApp.fetch(url, opciones);
    if (respuesta.getResponseCode() !== 200) {
        throw new Error(`Error al exportar a PDF. Código de respuesta: ${respuesta.getResponseCode()}`);
    }
    const blob = respuesta.getBlob().setName(codigo + '.pdf');
    carpetaDestino.createFile(blob);
    const sha512 = calcularSHA512(blob.getBytes());
    const fileSha512 = carpetaDestino.createFile(codigo + ".hexhash", sha512 + 'h');
    if (correo && correo.includes('@')) {
      const templ = HtmlService.createTemplateFromFile('mail');
      templ.codigo = codigo;
      templ.sha512 = sha512;
      const message = templ.evaluate().getContent();
      MailApp.sendEmail({
        name: 'Oposiciones Ayuntamiento El Campello',
        to: correo,
        subject: "Copia ejercicio Proceso Selectivo. NO RESPONDER",
        htmlBody: message,
        attachments: [blob, fileSha512]
      });
      Logger.log("Correo de confirmación enviado a " + correo);
    }
  } catch(e) {
    Logger.log(`Error crítico en exportarYGuardarEntregables para el código ${codigo}. Error: ${e.toString()}`);
    throw new Error(`No se pudo enviar el correo de confirmación. Error: ${e.message}`);
  }
}

function calcularSHA512(bytes) {
  if (bytes.length > 26214400) {
      return "Archivo demasiado grande para cálculo de SHA512 (límite 25MB). Tamaño: " + bytes.length;
  }
  try {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_512, bytes);
    return digest.map(byte => {
      const hex = (byte & 0xFF).toString(16);
      return hex.length === 1 ? '0' + hex : hex;
    }).join('').toUpperCase();
  } catch (e) {
    Logger.log(`Error al calcular SHA512: ${e}`);
    return "Error al calcular hash";
  }
}
