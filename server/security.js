import crypto from 'node:crypto';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import { z } from 'zod';

const contactSchema = z.object({
  name: z.string().min(2).max(80),
  email: z.string().email().max(254),
  message: z.string().min(10).max(2000),
  company: z.string().max(200).optional()
});

export function buildSecurityMiddleware({ allowedOrigin, isProd }) {
  const googleOrigins = ['https://accounts.google.com', 'https://www.gstatic.com'];
  const cspDirectives = {
    defaultSrc: ["'self'"],
    scriptSrc: ["'self'", ...googleOrigins],
    styleSrc: ["'self'"],
    imgSrc: ["'self'", 'data:', 'blob:', ...googleOrigins, 'https://lh3.googleusercontent.com'],
    fontSrc: ["'self'"],
    connectSrc: ["'self'", ...googleOrigins],
    objectSrc: ["'none'"],
    baseUri: ["'self'"],
    frameAncestors: ["'none'"],
    frameSrc: ["'self'", 'https://accounts.google.com'],
    formAction: ["'self'"],
    scriptSrcAttr: ["'none'"],
    styleSrcAttr: ["'unsafe-inline'"]
  };

  if (isProd) {
    cspDirectives.upgradeInsecureRequests = [];
  } else {
    cspDirectives.scriptSrc.push("'unsafe-eval'");
    cspDirectives.styleSrc.push("'unsafe-inline'");
    cspDirectives.connectSrc.push('ws:', 'wss:');
  }

  const helmetMiddleware = helmet({
    contentSecurityPolicy: { directives: cspDirectives },
    crossOriginEmbedderPolicy: false,
    crossOriginResourcePolicy: { policy: 'same-site' },
    hsts: isProd
  });

  const corsMiddleware = cors({
    origin: (origin, callback) => {
      if (!origin) {
        return callback(null, true);
      }
      if (origin === allowedOrigin) {
        return callback(null, true);
      }
      return callback(new Error('CORS_BLOCKED'));
    },
    methods: ['GET', 'POST'],
    allowedHeaders: ['Content-Type', 'X-CSRF-Token'],
    maxAge: 600
  });

  return { helmetMiddleware, corsMiddleware };
}

export function createRateLimiters({
  apiWindowMs = 10 * 60 * 1000,
  apiMax = 120,
  contactWindowMs = 10 * 60 * 1000,
  contactMax = 5
} = {}) {
  const apiLimiter = rateLimit({
    windowMs: apiWindowMs,
    max: apiMax,
    standardHeaders: true,
    legacyHeaders: false,
    handler: (_req, res) => {
      res.status(429).json({ error: 'rate_limited' });
    }
  });

  const contactLimiter = rateLimit({
    windowMs: contactWindowMs,
    max: contactMax,
    standardHeaders: true,
    legacyHeaders: false,
    handler: (_req, res) => {
      res.status(429).json({ error: 'rate_limited' });
    }
  });

  return { apiLimiter, contactLimiter };
}

export function createCsrf({ secret, ttlMs = 30 * 60 * 1000 }) {
  const issueToken = () => {
    const issuedAt = Date.now();
    const nonce = crypto.randomBytes(16).toString('hex');
    const payload = `${issuedAt}.${nonce}`;
    const signature = crypto.createHmac('sha256', secret).update(payload).digest('base64url');
    return {
      token: `${payload}.${signature}`,
      expiresAt: issuedAt + ttlMs
    };
  };

  const verifyToken = (token) => {
    if (!token || typeof token !== 'string') {
      return false;
    }
    const parts = token.split('.');
    if (parts.length !== 3) {
      return false;
    }
    const [issuedAtRaw, nonce, signature] = parts;
    const issuedAt = Number(issuedAtRaw);
    if (!Number.isFinite(issuedAt)) {
      return false;
    }
    if (issuedAt > Date.now() + 5 * 60 * 1000) {
      return false;
    }
    if (Date.now() - issuedAt > ttlMs) {
      return false;
    }
    if (!nonce || nonce.length < 16) {
      return false;
    }
    const payload = `${issuedAt}.${nonce}`;
    const expected = crypto.createHmac('sha256', secret).update(payload).digest('base64url');
    if (expected.length !== signature.length) {
      return false;
    }
    return crypto.timingSafeEqual(Buffer.from(expected), Buffer.from(signature));
  };

  return { issueToken, verifyToken, ttlMs };
}

export function enforceSameOrigin(allowedOrigin) {
  return (req, res, next) => {
    const origin = req.get('origin');
    const referer = req.get('referer');
    if (origin && origin !== allowedOrigin) {
      return res.status(403).json({ error: 'forbidden' });
    }
    if (!origin && referer && !referer.startsWith(allowedOrigin)) {
      return res.status(403).json({ error: 'forbidden' });
    }
    return next();
  };
}

export function verifyCsrfToken(verifyToken) {
  return (req, res, next) => {
    const token = req.get('x-csrf-token');
    if (!verifyToken(token)) {
      return res.status(403).json({ error: 'csrf_invalid' });
    }
    return next();
  };
}

export function validateContact(payload) {
  const parsed = contactSchema.safeParse(payload);
  if (!parsed.success) {
    return { ok: false, error: 'invalid_input' };
  }

  const data = parsed.data;
  const isSpam = Boolean(data.company && data.company.trim().length > 0);

  return {
    ok: true,
    isSpam,
    data: {
      name: sanitizeText(data.name, 80),
      email: sanitizeEmail(data.email),
      message: sanitizeMessage(data.message, 2000)
    }
  };
}

export function sanitizeText(value, maxLength) {
  const cleaned = value
    .normalize('NFKC')
    .replace(/\r\n/g, '\n')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim();
  if (maxLength && cleaned.length > maxLength) {
    return cleaned.slice(0, maxLength);
  }
  return cleaned;
}

export function sanitizeEmail(value) {
  const cleaned = sanitizeText(value, 254);
  return cleaned.toLowerCase();
}

export function sanitizeMessage(value, maxLength) {
  const cleaned = value
    .normalize('NFKC')
    .replace(/\r\n/g, '\n')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim();
  if (maxLength && cleaned.length > maxLength) {
    return cleaned.slice(0, maxLength);
  }
  return cleaned;
}
