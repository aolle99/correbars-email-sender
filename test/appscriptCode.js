// ─────────────────────────────────────────────
//  CONFIGURACIÓ GENERAL
// ─────────────────────────────────────────────
const CONFIG = {
    VERCEL_API_URL: 'https://correbars-email-sender.vercel.app/api/send',
    EMAIL_SUBJECT_PREFIX: '🍻 CORREBARS 2026 🍻',
    SENDER_NAME: 'Organització Correbars 2026',
    BASE_PRICE: 25,
    BOCATA_EXTRA: 3,
    NO_BOCATA_VALUE: '❌No, no en vull',
    LOG_SHEET_NAME: 'Envíos',
    ID_COLUMN_HEADER: 'ID Inscripció',
    STATUS_COLUMN_HEADER: 'Estat Enviament',
    RETRY_TRIGGER_COLUMN: 13,   // columna que activa el reintent manual
    LOCK_TIMEOUT_MS: 30000,
};

// ─────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────

/** Retorna les propietats de l'script de forma segura. */
function getProps() {
    return PropertiesService.getScriptProperties();
}

/** Genera un nou ID correlatiu amb bloqueig concurrent. */
function generateRegistrationId() {
    const lock = LockService.getScriptLock();
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    try {
        const props = getProps();
        const lastId = parseInt(props.getProperty('LAST_REG_ID') || '0', 10);
        const newId  = lastId + 1;
        props.setProperty('LAST_REG_ID', String(newId));
        return newId;
    } finally {
        lock.releaseLock();
    }
}

/** Calcula el preu final. */
function calcPrice(bocata) {
    return CONFIG.BASE_PRICE + (bocata !== CONFIG.NO_BOCATA_VALUE ? CONFIG.BOCATA_EXTRA : 0);
}

/** Troba o crea una columna pel header indicat. Retorna l'índex (base-1). */
function getOrCreateColumn(sheet, headerName) {
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
        // Full buit: crear la columna directament
        sheet.getRange(1, 1).setValue(headerName);
        return 1;
    }
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idx     = headers.indexOf(headerName);
    if (idx !== -1) return idx + 1; // base-1
    // No existeix: afegir al final
    const newCol = lastCol + 1;
    sheet.getRange(1, newCol).setValue(headerName);
    return newCol;
}

/** Formata una data en format local. */
function formatDate(dateValue) {
    const d = new Date(dateValue);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

/**
 * Renderitza la plantilla HTML i retorna el contingut.
 * Tots els valors es passen com a strings nets.
 */
function renderTemplate(data) {
    const tpl      = HtmlService.createTemplateFromFile('TemplateReserva');
    tpl.regId      = data.regId;
    tpl.nombre     = data.nombre;
    tpl.talla      = data.talla;
    tpl.alergias   = data.alergias   || 'Cap';
    tpl.bocata     = data.bocata;
    tpl.dni        = data.dni;
    tpl.data       = data.dataFormatada;
    tpl.preuFinal  = data.preuFinal;
    return tpl.evaluate().getContent();
}

/**
 * Envia el correu via Vercel (amb signatura HMAC opcional).
 * Llança una excepció si la resposta no és 200.
 */
function sendViaVercel(to, subject, htmlBody) {
    const props    = getProps();
    const API_KEY  = props.getProperty('API_KEY');
    const HMAC_SEC = props.getProperty('HMAC_SECRET');

    let signature = null;
    if (HMAC_SEC) {
        const raw        = to + subject + htmlBody;
        const rawBytes   = Utilities.newBlob(raw).getBytes();
        const secretBytes = Utilities.newBlob(HMAC_SEC).getBytes();
        const hmacBytes  = Utilities.computeHmacSha256Signature(rawBytes, secretBytes);
        signature        = hmacBytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
    }

    const payload = { to, subject, html: htmlBody };
    if (signature) payload.signature = signature;

    const options = {
        method:             'post',
        contentType:        'application/json',
        headers:            { 'x-api-key': API_KEY || '' },
        payload:            JSON.stringify(payload),
        muteHttpExceptions: true,
    };

    const resp = UrlFetchApp.fetch(CONFIG.VERCEL_API_URL, options);
    const code = resp.getResponseCode();
    if (code !== 200) {
        throw new Error(`Vercel resposta ${code}: ${resp.getContentText()}`);
    }
}

/**
 * Envia el correu amb fallback a GmailApp/MailApp.
 * Retorna true si s'ha enviat correctament per qualsevol via.
 */
function sendEmailWithFallback(to, subject, htmlBody) {
    // Intent 1: Vercel
    try {
        sendViaVercel(to, subject, htmlBody);
        Logger.log('✅ Enviat via Vercel a: ' + to);
        return true;
    } catch (errVercel) {
        Logger.log('⚠️ Vercel fallat: ' + errVercel.message);
    }

    // Intent 2: GmailApp
    try {
        GmailApp.sendEmail(to, subject, '', {
            name:     CONFIG.SENDER_NAME,
            htmlBody: htmlBody,
        });
        Logger.log('✅ Enviat via GmailApp a: ' + to);
        return true;
    } catch (errGmail) {
        Logger.log('⚠️ GmailApp fallat: ' + errGmail.message);
    }

    // Intent 3: MailApp (quota diferent)
    try {
        MailApp.sendEmail({
            to:       to,
            name:     CONFIG.SENDER_NAME,
            subject:  subject,
            htmlBody: htmlBody,
        });
        Logger.log('✅ Enviat via MailApp a: ' + to);
        return true;
    } catch (errMail) {
        Logger.log('❌ MailApp també fallat: ' + errMail.message);
    }

    return false;
}

/** Escriu una entrada al full de registre d'enviaments. */
function logToEnviosSheet(email, regId, preuFinal, status) {
    try {
        const ss          = SpreadsheetApp.getActiveSpreadsheet();
        const enviosSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
        if (!enviosSheet) {
            Logger.log('⚠️ No s\'ha trobat el full "' + CONFIG.LOG_SHEET_NAME + '"');
            return;
        }
        enviosSheet.appendRow([
            new Date(),
            email,
            regId,
            preuFinal + '€',
            status,
        ]);
    } catch (err) {
        Logger.log('⚠️ Error escrivint al full Envíos: ' + err.message);
    }
}

// ─────────────────────────────────────────────
//  TRIGGER PRINCIPAL: onFormSubmit
// ─────────────────────────────────────────────
function onFormSubmit(e) {
    // ── 1) Genera ID únic ─────────────────────
    const newId = generateRegistrationId();

    // ── 2) Extreu dades del formulari ─────────
    const r          = e.namedValues;
    const emailDest  = (r['Adreça electrònica']          || [''])[0].trim();
    const nombre     = (r['🙍‍♀️ Nom i Cognoms']         || [''])[0].trim();
    const talla      = (r['👕 Talla de samarreta']        || [''])[0].trim();
    const alergias   = (r['🍃Al·lèrgies o intoleràncies'] || ['Cap'])[0].trim() || 'Cap';
    const bocata     = (r['🥪Entrepà']                    || [''])[0].trim();
    const dni        = (r['🪪 Número DNI']                || [''])[0].trim();
    const timestamp  = (r['Marca de temps']               || [new Date().toString()])[0];

    const preuFinal    = calcPrice(bocata);
    const dataFormatada = formatDate(timestamp);

    // Validació bàsica
    if (!emailDest) {
        Logger.log('❌ onFormSubmit: email buit, abort.');
        return;
    }

    // ── 3) Escriu l'ID i l'estat a la fila ────
    const range = e.range;
    const row   = range.getRow();
    const sheet = range.getSheet();

    const idCol     = getOrCreateColumn(sheet, CONFIG.ID_COLUMN_HEADER);
    const statusCol = getOrCreateColumn(sheet, CONFIG.STATUS_COLUMN_HEADER);
    sheet.getRange(row, idCol).setValue(newId);
    sheet.getRange(row, statusCol).setValue('PENDENT');

    // ── 4) Renderitza la plantilla ────────────
    let htmlBody;
    try {
        htmlBody = renderTemplate({ regId: newId, nombre, talla, alergias, bocata, dni, dataFormatada, preuFinal });
    } catch (errTpl) {
        Logger.log('❌ Error renderitzant plantilla: ' + errTpl.message);
        sheet.getRange(row, statusCol).setValue('ERROR_PLANTILLA');
        return;
    }

    const subject = `${CONFIG.EMAIL_SUBJECT_PREFIX} #${newId}`;

    // ── 5) Envia el correu ───────────────────
    const envioExitoso = sendEmailWithFallback(emailDest, subject, htmlBody);
    const finalStatus  = envioExitoso ? 'OK' : 'ERROR_ENVIAMENT';
    sheet.getRange(row, statusCol).setValue(finalStatus);

    // ── 6) Log al full Envíos ────────────────
    logToEnviosSheet(emailDest, newId, preuFinal, envioExitoso ? 'OK' : 'KO');

    Logger.log(`onFormSubmit finalitzat — ID: ${newId} | Email: ${emailDest} | Estat: ${finalStatus}`);
}

// ─────────────────────────────────────────────
//  TRIGGER: retryFailedEmail (edició manual)
//  Marca la cel·la de la columna STATUS com a
//  "RETRY" per tornar a enviar el correu.
// ─────────────────────────────────────────────
function retryFailedEmail(e) {
    const sheet = e.source.getActiveSheet();
    const range = e.range;

    // Només actua si s'ha editat la columna de reintent i el valor és true/RETRY
    const triggerValue = String(range.getValue()).toUpperCase();
    if (range.getColumn() !== CONFIG.RETRY_TRIGGER_COLUMN) return;
    if (triggerValue !== 'TRUE' && triggerValue !== 'RETRY') return;

    const row = range.getRow();
    if (row <= 1) return; // ignora la capçalera

    // Llegeix les dades de la fila
    const emailDest = sheet.getRange(row, 2).getValue();
    const nombre    = sheet.getRange(row, 3).getValue();
    const dni       = sheet.getRange(row, 4).getValue();
    const talla     = sheet.getRange(row, 5).getValue();
    const bocata    = sheet.getRange(row, 6).getValue();
    const alergias  = sheet.getRange(row, 8).getValue() || 'Cap';
    const fecha     = new Date(sheet.getRange(row, 1).getValue());
    const newId     = sheet.getRange(row, 12).getValue();

    if (!emailDest || !newId) {
        Logger.log('⚠️ retryFailedEmail: dades incompletes a la fila ' + row);
        return;
    }

    const preuFinal     = calcPrice(bocata);
    const dataFormatada = formatDate(fecha);

    let htmlBody;
    try {
        htmlBody = renderTemplate({ regId: newId, nombre, talla, alergias, bocata, dni, dataFormatada, preuFinal });
    } catch (errTpl) {
        Logger.log('❌ retryFailedEmail error plantilla: ' + errTpl.message);
        return;
    }

    const subject        = `${CONFIG.EMAIL_SUBJECT_PREFIX} #${newId}`;
    const envioExitoso   = sendEmailWithFallback(emailDest, subject, htmlBody);
    const retryStatus    = envioExitoso ? 'RETRY_OK' : 'RETRY_KO';

    // Actualitza l'estat a la columna STATUS
    const statusCol = getOrCreateColumn(sheet, CONFIG.STATUS_COLUMN_HEADER);
    sheet.getRange(row, statusCol).setValue(retryStatus);

    // Log
    logToEnviosSheet(emailDest, newId, preuFinal, retryStatus);
    Logger.log(`retryFailedEmail — ID: ${newId} | Email: ${emailDest} | Estat: ${retryStatus}`);
}

// ─────────────────────────────────────────────
//  UTILITAT: reenviar tots els KO en batch
//  Executa manualment des de l'editor si cal.
// ─────────────────────────────────────────────
function retryAllFailed() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0]; // primera fulla (respostes)
    const data  = sheet.getDataRange().getValues();
    const heads = data[0];

    const statusColIdx = heads.indexOf(CONFIG.STATUS_COLUMN_HEADER);
    const idColIdx     = heads.indexOf(CONFIG.ID_COLUMN_HEADER);
    if (statusColIdx === -1 || idColIdx === -1) {
        Logger.log('⚠️ No s\'han trobat les columnes de control.');
        return;
    }

    let retried = 0;
    for (let i = 1; i < data.length; i++) {
        const status = String(data[i][statusColIdx]).toUpperCase();
        if (status !== 'KO' && status !== 'ERROR_ENVIAMENT' && status !== 'RETRY_KO') continue;

        const row        = i + 1;
        const emailDest  = data[i][1];
        const nombre     = data[i][2];
        const dni        = data[i][3];
        const talla      = data[i][4];
        const bocata     = data[i][5];
        const alergias   = data[i][7] || 'Cap';
        const fecha      = new Date(data[i][0]);
        const regId      = data[i][idColIdx];

        if (!emailDest || !regId) continue;

        const preuFinal     = calcPrice(bocata);
        const dataFormatada = formatDate(fecha);

        let htmlBody;
        try {
            htmlBody = renderTemplate({ regId, nombre, talla, alergias, bocata, dni, dataFormatada, preuFinal });
        } catch (err) {
            Logger.log(`❌ Fila ${row}: error plantilla — ${err.message}`);
            continue;
        }

        const subject      = `${CONFIG.EMAIL_SUBJECT_PREFIX} #${regId}`;
        const ok           = sendEmailWithFallback(emailDest, subject, htmlBody);
        const newStatus    = ok ? 'RETRY_OK' : 'RETRY_KO';
        sheet.getRange(row, statusColIdx + 1).setValue(newStatus);
        logToEnviosSheet(emailDest, regId, preuFinal, newStatus);
        retried++;

        // Pausa per evitar límits de quota
        Utilities.sleep(1000);
    }

    Logger.log(`retryAllFailed completat: ${retried} registres reintentats.`);
}