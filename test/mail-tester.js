import 'dotenv/config';
import crypto from 'crypto';

const API_ENDPOINT = process.env.API_ENDPOINT;
const API_KEY     = process.env.API_KEY;
const HMAC_SECRET = process.env.HMAC_SECRET;
const EMAIL_TO    = process.env.EMAIL_TO;

// --- Validació de variables d'entorn ---
const missing = ['API_ENDPOINT', 'API_KEY', 'EMAIL_TO'].filter(k => !process.env[k]);
if (missing.length) {
    console.error(`❌ Falten variables al .env: ${missing.join(', ')}`);
    process.exit(1);
}

function generateSignature(to, subject, html) {
    if (!HMAC_SECRET) return null;
    const hmac = crypto.createHmac('sha256', HMAC_SECRET);
    hmac.update(to + subject + html);
    return hmac.digest('hex');
}

// --- Contingut del correu (similar al real de producció) ---
const entradaNum = '#TEST-001';
const subject = `Confirmació d'inscripció ${entradaNum} - Correbars Esparreguera`;
const html = `<!DOCTYPE html>
<html lang="ca">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Confirmació Correbars</title>
</head>
<body style="margin:0; padding:0; background-color:#f4f4f4; font-family: Arial, sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f4f4f4; padding: 20px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">

          <!-- Capçalera -->
          <tr>
            <td style="background-color:#1a1a2e; padding: 30px; text-align:center;">
              <h1 style="color:#d4af37; margin:0; font-size:26px;">Correbars Esparreguera</h1>
              <p style="color:#cccccc; margin:8px 0 0 0; font-size:14px;">Confirmació d'inscripció</p>
            </td>
          </tr>

          <!-- Cos -->
          <tr>
            <td style="padding: 30px;">
              <p style="color:#333333; font-size:16px;">Hola!</p>
              <p style="color:#555555; font-size:15px; line-height:1.6;">
                Et confirmem que la teva inscripció als Correbars d'Esparreguera s'ha registrat correctament.
              </p>

              <!-- Taula de detalls -->
              <table width="100%" cellpadding="0" cellspacing="0" style="margin: 24px 0; border-collapse: collapse;">
                <tr style="background-color:#f9f9f9;">
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; font-weight:bold; color:#333; width:45%;">Número d'entrada</td>
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; color:#555;">#TEST-001</td>
                </tr>
                <tr>
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; font-weight:bold; color:#333;">Nom</td>
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; color:#555;">Participant de prova</td>
                </tr>
                <tr style="background-color:#f9f9f9;">
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; font-weight:bold; color:#333;">Preu</td>
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; color:#555;">25,00 €</td>
                </tr>
                <tr>
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; font-weight:bold; color:#333;">Data de l'event</td>
                  <td style="padding: 12px 15px; border: 1px solid #e0e0e0; color:#555;">Per confirmar</td>
                </tr>
              </table>

              <p style="color:#555555; font-size:15px; line-height:1.6;">
                Si tens qualsevol dubte, pots respondre directament a aquest correu.
              </p>
              <p style="color:#555555; font-size:15px;">Fins aviat!</p>
            </td>
          </tr>

          <!-- Peu de pàgina -->
          <tr>
            <td style="background-color:#f0f0f0; padding: 20px 30px; text-align:center; border-top: 1px solid #e0e0e0;">
              <p style="color:#888888; font-size:12px; margin:0; line-height:1.6;">
                Organització Correbars Esparreguera<br>
                Has rebut aquest correu perquè et vas inscriure a l'event.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

async function sendToMailTester() {
    console.log('╔════════════════════════════════════════╗');
    console.log('║      TEST DE DELIVERABILITY            ║');
    console.log('║      mail-tester.com                   ║');
    console.log('╚════════════════════════════════════════╝\n');
    console.log(`📡 Endpoint : ${API_ENDPOINT}`);
    console.log(`📧 Destinatari: ${EMAIL_TO}`);
    console.log(`📝 Assumpte  : ${subject}`);
    console.log(`🔐 HMAC      : ${HMAC_SECRET ? 'activat' : 'desactivat'}\n`);

    const payload = {
        to: EMAIL_TO,
        subject,
        html,
        ...(HMAC_SECRET && { signature: generateSignature(EMAIL_TO, subject, html) })
    };

    try {
        console.log('📤 Enviant...');
        const response = await fetch(API_ENDPOINT, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': API_KEY
            },
            body: JSON.stringify(payload)
        });

        const result = await response.json();

        if (response.ok) {
            console.log(`\n✅ Correu enviat correctament!`);
            console.log(`   messageId : ${result.messageId}`);
            console.log(`   timestamp : ${result.timestamp}`);
            console.log('\n══════════════════════════════════════════');
            console.log('🔍 Consulta la puntuació a:');

            // Extreu la URL de mail-tester a partir del correu de destí
            // Format: test-XXXXX@srv1.mail-tester.com  →  https://www.mail-tester.com/test-XXXXX
            const match = EMAIL_TO.match(/^(test-[^@]+)@/);
            if (match) {
                console.log(`   https://www.mail-tester.com/${match[1]}`);
            } else {
                console.log(`   https://www.mail-tester.com`);
            }
            console.log('══════════════════════════════════════════\n');
        } else {
            console.error(`\n❌ Error en l'enviament (HTTP ${response.status})`);
            console.error(`   ${result.error || JSON.stringify(result)}`);
            process.exit(1);
        }
    } catch (err) {
        console.error(`\n❌ Error de connexió: ${err.message}`);
        console.error('   Comprova que API_ENDPOINT és correcte i Vercel està desplegat.');
        process.exit(1);
    }
}

sendToMailTester();
