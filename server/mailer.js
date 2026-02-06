import nodemailer from 'nodemailer';
import { sanitizeEmail, sanitizeMessage, sanitizeText } from './security.js';

export function createMailer(env) {
  const host = env.SMTP_HOST || '';
  const port = Number(env.SMTP_PORT || '587');
  const secure = env.SMTP_SECURE === 'true';
  const user = env.SMTP_USER || '';
  const pass = env.SMTP_PASS || '';
  const contactTo = env.CONTACT_TO || 'infinitydevbr@gmail.com';
  const contactFrom = env.CONTACT_FROM || 'Infinite Dev <no-reply@infinite.dev>';

  const isConfigured = Boolean(host && contactTo);
  let transporter = null;

  if (isConfigured) {
    transporter = nodemailer.createTransport({
      host,
      port,
      secure,
      auth: user && pass ? { user, pass } : undefined
    });
  }

  const sendContact = async ({ name, email, message }) => {
    const safeName = sanitizeText(name, 80);
    const safeEmail = sanitizeEmail(email);
    const safeMessage = sanitizeMessage(message, 2000);

    if (!transporter) {
      console.info('contact_message_no_smtp', {
        name: safeName,
        email: safeEmail,
        message: safeMessage.slice(0, 500)
      });
      return { delivered: false };
    }

    await transporter.sendMail({
      to: contactTo,
      from: sanitizeText(contactFrom, 120),
      replyTo: { name: safeName, address: safeEmail },
      subject: `Contato - ${safeName}`,
      text: safeMessage
    });

    return { delivered: true };
  };

  return { sendContact, isConfigured };
}
