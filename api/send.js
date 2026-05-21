import nodemailer from 'nodemailer';
import rateLimit from 'express-rate-limit';
import crypto from 'crypto';

// Función para validar email
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

// Función para sanitizar HTML básico
function sanitizeHtml(html) {
    if (typeof html !== 'string') return '';
    return html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
        .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '');
}

const API_KEY = process.env.API_KEY;
const OVH_USER = process.env.OVH_USER;
const OVH_PASS = process.env.OVH_PASS;
const HMAC_SECRET = process.env.HMAC_SECRET;

function getClientIP(req) {
    return req.headers['x-forwarded-for']?.split(',')[0]?.trim() ||
        req.headers['x-real-ip'] ||
        req.connection?.remoteAddress ||
        req.socket?.remoteAddress ||
        req.ip ||
        'unknown';
}

const limiter = rateLimit({
    windowMs: 60 * 1000,
    max: 30,
    standardHeaders: true,
    legacyHeaders: false,
    message: { error: 'Too many requests, please try again later.' },
    keyGenerator: (req) => {
        const clientIP = getClientIP(req);
        const apiKey = req.headers['x-api-key'] || 'no-key';
        return `${clientIP}-${apiKey}`;
    },
    skip: (req) => {
        return false;
    }
});

export default async function handler(req, res) {
    try {
        // Configurar headers CORS y de respuesta
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-api-key');

        if (req.method === 'OPTIONS') {
            return res.status(200).json({ ok: true });
        }

        if (req.method !== 'POST') {
            return res.status(405).json({ error: 'Method not allowed' });
        }

        // Verificar variables de entorno críticas
        if (!API_KEY || !OVH_USER || !OVH_PASS) {
            console.error('Missing environment variables');
            return res.status(500).json({ error: 'Server configuration error' });
        }

        const incomingKey = req.headers['x-api-key'];
        if (!incomingKey || incomingKey !== API_KEY) {
            return res.status(401).json({ error: 'Unauthorized' });
        }

        // Rate limiting
        try {
            await new Promise((resolve, reject) => {
                limiter(req, res, (err) => err ? reject(err) : resolve());
            });
        } catch (rateLimitError) {
            console.error('Rate limit error:', rateLimitError);
        }

        let body;
        try {
            body = typeof req.body === 'object' ? req.body : JSON.parse(req.body);
        } catch (e) {
            console.error('JSON parse error:', e);
            return res.status(400).json({ error: 'Invalid JSON format' });
        }

        const { to, subject, html, signature, attachments } = body;

        // Validaciones
        if (!to || !subject || !html) {
            return res.status(400).json({
                error: 'Missing required fields: to, subject, html'
            });
        }

        if (!isValidEmail(to)) {
            return res.status(400).json({ error: 'Invalid email address' });
        }

        if (subject.length > 200) {
            return res.status(400).json({ error: 'Subject too long (max 200 chars)' });
        }

        // Validar attachments si existen con límites más estrictos
        let totalAttachmentSize = 0;
        if (attachments && Array.isArray(attachments)) {
            for (let i = 0; i < attachments.length; i++) {
                const att = attachments[i];
                if (!att.filename || !att.content || !att.contentType) {
                    return res.status(400).json({ 
                        error: `Invalid attachment format at index ${i}` 
                    });
                }
                
                // Verificar extensiones permitidas
                const allowedExtensions = ['.pdf', '.txt', '.jpg', '.jpeg', '.png', '.gif', '.doc', '.docx', '.xls', '.xlsx', '.csv'];
                const fileExt = att.filename.toLowerCase().substring(att.filename.lastIndexOf('.'));
                if (!allowedExtensions.includes(fileExt)) {
                    return res.status(400).json({ 
                        error: `File type not allowed: ${fileExt}. Allowed: ${allowedExtensions.join(', ')}` 
                    });
                }
                
                // Calcular tamaño del archivo (base64 -> bytes reales)
                const fileSize = Math.ceil(att.content.length * 0.75);
                totalAttachmentSize += fileSize;
                
                // Límite por archivo individual: 10MB
                if (fileSize > 10 * 1024 * 1024) {
                    return res.status(400).json({ 
                        error: `Attachment "${att.filename}" too large (max 10MB per file)` 
                    });
                }
            }

            // Límite total: 15MB (más conservador)
            if (totalAttachmentSize > 15 * 1024 * 1024) {
                return res.status(400).json({ 
                    error: `Total attachments too large (max 15MB total). Current size: ${Math.round(totalAttachmentSize / 1024 / 1024)}MB` 
                });
            }
        }

        const sanitizedHtml = sanitizeHtml(html);

        // HMAC validation solo en producción
        const isProduction = process.env.NODE_ENV === 'production';
        if (HMAC_SECRET && isProduction) {
            if (!signature) {
                return res.status(403).json({ error: 'Signature required' });
            }

            const hmac = crypto.createHmac('sha256', HMAC_SECRET);
            hmac.update(to + subject + html);
            const expected = hmac.digest('hex');

            if (signature !== expected) {
                return res.status(403).json({ error: 'Invalid signature' });
            }
        }

        // Configurar transporter con timeouts más largos para adjuntos
        const transporter = nodemailer.createTransport({
            host: 'ssl0.ovh.net',
            port: 465,
            secure: true,
            auth: {
                user: OVH_USER,
                pass: OVH_PASS
            },
            connectionTimeout: 30000,  // 30 segundos
            greetingTimeout: 10000,    // 10 segundos
            socketTimeout: 60000,      // 60 segundos para adjuntos grandes
            pool: true,
            maxConnections: 3,         // Menos conexiones concurrentes
            maxMessages: 50,
            rateDelta: 30000,          // Más tiempo entre emails
            rateLimit: 3,              // Menos emails por período
        });

        // Preparar attachments para nodemailer con validación adicional
        let emailAttachments = [];
        if (attachments && Array.isArray(attachments)) {
            emailAttachments = attachments.map((att, index) => {
                try {
                    // Verificar que el base64 sea válido
                    const buffer = Buffer.from(att.content, 'base64');
                    if (buffer.length === 0) {
                        throw new Error(`Empty attachment content at index ${index}`);
                    }
                    
                    return {
                        filename: att.filename,
                        content: buffer,
                        contentType: att.contentType,
                        encoding: 'base64',
                        // Headers adicionales para evitar problemas
                        cid: `attachment_${index}@correbars.local`
                    };
                } catch (error) {
                    throw new Error(`Invalid base64 content in attachment ${att.filename}: ${error.message}`);
                }
            });
        }

        // Crear texto plano más limpio para evitar filtros de spam
        const textContent = html
            .replace(/<[^>]*>/g, '')
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/\s+/g, ' ')
            .trim();

        const mailOptions = {
            from: {
                name: 'Correbars Esparreguera',
                address: OVH_USER
            },
            replyTo: {
                name: 'Correbars Esparreguera',
                address: OVH_USER
            },
            to,
            subject,
            text: textContent,
            html: sanitizedHtml,
            attachments: emailAttachments,
            date: new Date(),
            encoding: 'utf-8',
            disableFileAccess: true,
            disableUrlAccess: true
        };

        console.log(`Sending email to ${to} with ${emailAttachments.length} attachments (${Math.round(totalAttachmentSize / 1024)}KB total)`);

        // Enviar email con timeout extendido
        const emailPromise = transporter.sendMail(mailOptions);
        const timeoutPromise = new Promise((_, reject) => 
            setTimeout(() => reject(new Error('Email sending timeout')), 90000) // 90 segundos
        );

        const info = await Promise.race([emailPromise, timeoutPromise]);

        console.log('Email sent successfully:', {
            messageId: info.messageId,
            response: info.response,
            attachmentCount: emailAttachments.length
        });

        return res.status(200).json({
            ok: true,
            messageId: info.messageId,
            timestamp: new Date().toISOString(),
            attachmentCount: emailAttachments.length,
            totalSize: `${Math.round(totalAttachmentSize / 1024)}KB`
        });

    } catch (err) {
        console.error('API Error:', {
            message: err.message,
            stack: err.stack,
            timestamp: new Date().toISOString(),
            hasAttachments: !!(req.body?.attachments?.length)
        });

        // Error específico para problemas de adjuntos
        if (err.message.includes('attachment') || err.message.includes('base64') || err.message.includes('timeout')) {
            return res.status(400).json({
                ok: false,
                error: 'Attachment processing error: ' + err.message,
                timestamp: new Date().toISOString()
            });
        }

        return res.status(500).json({
            ok: false,
            error: 'Internal server error',
            timestamp: new Date().toISOString()
        });
    }
}