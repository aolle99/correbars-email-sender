/**
 * Envía el correo usando tu endpoint de Vercel protegido.
 */
function sendViaVercel(to, subject, htmlBody, attachments = null) {
    // URL de tu deployment en Vercel
    const API_URL = 'https://correbars-email-sender.vercel.app/api/send';
    // Claves almacenadas en Script Properties
    const props = PropertiesService.getScriptProperties();
    const API_KEY = props.getProperty('API_KEY');
    const HMAC_SEC = props.getProperty('HMAC_SECRET');

    // Opcional: firma HMAC SHA256 de (to+subject+html)
    let signature = null;
    if (HMAC_SEC) {
        const raw = to + subject + htmlBody;
        const rawBytes = Utilities.newBlob(raw).getBytes();
        const secretBytes = Utilities.newBlob(HMAC_SEC).getBytes();

        const hmacBytes = Utilities.computeHmacSha256Signature(rawBytes, secretBytes);
        signature = hmacBytes
            .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
            .join('');
    }

    const payload = { to, subject, html: htmlBody };
    if (signature) payload.signature = signature;

    // Si hay adjuntos, convertirlos a base64 para enviar por API
    if (attachments && attachments.length > 0) {
        payload.attachments = attachments.map(attachment => ({
            filename: attachment.getName(),
            content: Utilities.base64Encode(attachment.getBytes()),
            contentType: attachment.getContentType()
        }));
    }

    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-api-key': API_KEY },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const resp = UrlFetchApp.fetch(API_URL, options);
    if (resp.getResponseCode() !== 200) {
        let errorMsg = resp.getContentText();
        try {
            const respJson = JSON.parse(errorMsg);
            if (respJson.error) errorMsg = respJson.error;
        } catch (_) {}
        throw new Error('Error en Vercel: ' + errorMsg);
    }
}

/**
 * Función principal para enviar entradas cuando se marca como pagado
 * Se dispara cuando se marca la casilla de "Estado de Pago" (columna 8)
 */
function enviarEntradas(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;

    // Verificar si el cambio fue en la columna 8 (Estado de Pago) y es una casilla marcada
    if (range.getColumn() == 10 && range.getValue() == true) {
        var row = range.getRow();

        // Extraer datos de la fila
        var regId = sheet.getRange(row, 1).getValue(); // Número de entrada (columna A)
        var nombre = sheet.getRange(row, 2).getValue(); // Nombre (columna B)
        var talla = sheet.getRange(row, 3).getValue(); // Talla (columna C)
        var correo = sheet.getRange(row, 4).getValue(); // Email (columna D)
        var dni = sheet.getRange(row, 5).getValue(); // DNI (columna E)
        var bocata = sheet.getRange(row, 8).getValue(); // Bocadillo (columna F)
        var alergias = sheet.getRange(row, 7).getValue() || "Cap"; // Alergias (columna G)
        var preuFinal = sheet.getRange(row, 9).getValue();
        var sending = sheet.getRange(row, 11).getValue(); // Estado sending (columna I)
        var estadoEnvio = sheet.getRange(row, 12).getValue(); // Estado envío (columna J)

        // Solo enviar si no se ha enviado ya y no está en proceso de envío
        if (estadoEnvio == false && sending == false) {
            // Marcar como "enviando"
            sheet.getRange(row, 11).setValue(true);

            try {
                // Renderizar la plantilla HTML de confirmación
                const tpl = HtmlService.createTemplateFromFile('TemplateConfirmation');
                tpl.regId = regId;
                tpl.nombre = nombre;
                tpl.talla = talla;
                tpl.alergias = alergias;
                tpl.bocata = bocata;
                tpl.dni = dni;
                tpl.data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
                tpl.preuFinal = preuFinal;

                const htmlBody = tpl.evaluate().getContent();

                // Crear PDF de la entrada
                var pdf = crearPDF(nombre, regId, bocata);

                var subject = `🍻 ENTRADA CORREBARS 2025 🍻 - ${nombre}`;

                // Intentar envío por Vercel
                let envioExitoso = false;
                let errorMsg = '';
                try {
                    sendViaVercel(correo, subject, htmlBody, [pdf]);
                    envioExitoso = true;
                    Logger.log('Enviado correctamente por Vercel para: ' + nombre);
                } catch (err) {
                    Logger.log('Falló envío Vercel: ' + err);
                    errorMsg = err.message;

                    // Fallback a Gmail si falla Vercel
                    try {
                        MailApp.sendEmail({
                            to: correo,
                            name: "Organització Correbars 2025",
                            subject: subject,
                            htmlBody: htmlBody,
                            attachments: [pdf]
                        });
                        envioExitoso = true;
                        errorMsg = '';
                        Logger.log('Enviado correctamente por Gmail para: ' + nombre);
                    } catch (gmailErr) {
                        Logger.log('También falló Gmail: ' + gmailErr);
                        errorMsg += ' | Gmail: ' + gmailErr.message;
                    }
                }

                // Actualizar estado del envío
                sheet.getRange(row, 12).setValue(envioExitoso);
                sheet.getRange(row, 11).setValue(false);

                // Registrar en hoja de "Envíos"
                registrarEnvio(correo, regId, preuFinal, envioExitoso ? 'OK' : 'KO', errorMsg);

            } catch (error) {
                Logger.log('Error general en envío: ' + error);
                sheet.getRange(row, 12).setValue(false);
                sheet.getRange(row, 11).setValue(false);

                // Registrar error en hoja de "Envíos"
                registrarEnvio(correo, regId, preuFinal, 'ERROR', error.message);
            }
        }
    }
}

/**
 * Función para crear el PDF de la entrada
 */
function crearPDF(nombre, numeroEntrada, bocata) {
    try {
        // ID del documento de Google Docs (plantilla de entrada)
        var templateId = '1P_V30EU7w5avGR8ijRnuB_aDQDdiANvuA-w31IAYTUE'; // Reemplaza con tu ID
        var copyId = DriveApp.getFileById(templateId).makeCopy().getId();
        var copyDoc = DocumentApp.openById(copyId);
        var copyBody = copyDoc.getBody();

        // Reemplazar los marcadores de posición en la plantilla
        copyBody.replaceText('{{Nombre}}', nombre);
        copyBody.replaceText('{{NumeroEntrada}}', numeroEntrada);
        copyBody.replaceText('{{Bocata}}', bocata);

        // Guardar y cerrar el documento
        copyDoc.saveAndClose();

        // Convertir el documento a PDF
        var pdf = DriveApp.getFileById(copyId).getAs('application/pdf');
        pdf.setName("Entrada Correbars 2025 - " + nombre + " #" + numeroEntrada + ".pdf");

        // Eliminar la copia del documento
        DriveApp.getFileById(copyId).setTrashed(true);

        return pdf;
    } catch (error) {
        Logger.log('Error creando PDF: ' + error);
        throw error;
    }
}

/**
 * Función para registrar envíos en la hoja de "Envíos"
 */
function registrarEnvio(correo, regId, precio, estado, mensaje) {
    try {
        const enviosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Envíos');
        if (enviosSheet) {
            enviosSheet.appendRow([
                new Date(),
                correo,
                regId,
                precio + '€',
                estado,
                mensaje || '',
            ]);
        }
    } catch (error) {
        Logger.log('Error registrando envío: ' + error);
    }
}
