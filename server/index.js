import 'dotenv/config';
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import express from 'express';
import { OAuth2Client } from 'google-auth-library';
import {
  buildSecurityMiddleware,
  createCsrf,
  createRateLimiters,
  enforceSameOrigin,
  validateContact,
  verifyCsrfToken
} from './security.js';
import { createMailer } from './mailer.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const docmakerHost = (process.env.DOCMAKER_HOST || 'docmaker.infinity.dev.br').toLowerCase();

let docmakerApp = null;
let docmakerLoadError = null;

const loadDocmakerApp = () => {
  if (docmakerApp || docmakerLoadError) {
    return docmakerApp;
  }

  try {
    const require = createRequire(import.meta.url);
    const docmakerModule = require(path.resolve(rootDir, 'docmaker', 'server.js'));
    docmakerApp = docmakerModule.app || docmakerModule;
  } catch (error) {
    docmakerLoadError = error;
    console.error('docmaker_load_failed', error);
  }

  return docmakerApp;
};

const isDocmakerHost = (req) => {
  const hostname = (req.hostname || '').toLowerCase();
  return hostname === docmakerHost;
};

const isProd = process.env.NODE_ENV === 'production' || process.argv.includes('--prod');
const port = Number(process.env.PORT || '3000');
const allowedOrigin = process.env.PUBLIC_ORIGIN || `http://localhost:${port}`;
const csrfSecret = process.env.CSRF_SECRET || (isProd ? '' : crypto.randomBytes(32).toString('hex'));

if (!csrfSecret) {
  throw new Error('CSRF_SECRET is required in production');
}

const app = express();
app.set('trust proxy', 1);
app.disable('x-powered-by');

app.use((req, res, next) => {
  if (!isDocmakerHost(req)) {
    return next();
  }

  const docmaker = loadDocmakerApp();
  if (!docmaker) {
    return res.status(503).send('Docmaker app unavailable');
  }

  return docmaker(req, res, (err) => {
    if (err) {
      return next(err);
    }
    if (!res.headersSent) {
      res.status(404).send('Not Found');
    }
  });
});

const { helmetMiddleware, corsMiddleware } = buildSecurityMiddleware({
  allowedOrigin,
  isProd
});

app.use(helmetMiddleware);

const { apiLimiter, contactLimiter } = createRateLimiters({
  apiWindowMs: Number(process.env.API_RATE_WINDOW_MS || '600000'),
  apiMax: Number(process.env.API_RATE_MAX || '120'),
  contactWindowMs: Number(process.env.CONTACT_RATE_WINDOW_MS || '600000'),
  contactMax: Number(process.env.CONTACT_RATE_MAX || '5')
});

app.use('/api', apiLimiter);
app.use('/api', corsMiddleware);
app.use('/api', (err, _req, res, next) => {
  if (err) {
    return res.status(403).json({ error: 'cors_blocked' });
  }
  return next();
});

app.use('/api', express.json({ limit: '10kb', type: 'application/json' }));
app.use('/api', express.urlencoded({ extended: false, limit: '10kb' }));

const { issueToken, verifyToken, ttlMs } = createCsrf({
  secret: csrfSecret,
  ttlMs: Number(process.env.CSRF_TTL_MS || '1800000')
});

const mailer = createMailer(process.env);
const googleClientId = process.env.GOOGLE_CLIENT_ID || '';
const googleClient = googleClientId ? new OAuth2Client(googleClientId) : null;

const verifyGoogleToken = async (idToken) => {
  if (!googleClient) {
    throw new Error('google_not_configured');
  }
  const ticket = await googleClient.verifyIdToken({
    idToken,
    audience: googleClientId
  });
  const payload = ticket.getPayload();
  if (!payload || !payload.email) {
    throw new Error('google_invalid');
  }
  return {
    email: payload.email,
    name: payload.name || payload.given_name || ''
  };
};

app.get('/api/csrf', enforceSameOrigin(allowedOrigin), (_req, res) => {
  const { token, expiresAt } = issueToken();
  res.set('Cache-Control', 'no-store');
  res.json({ token, expiresAt, ttlMs });
});

app.post(
  '/api/contact',
  enforceSameOrigin(allowedOrigin),
  verifyCsrfToken(verifyToken),
  contactLimiter,
  async (req, res) => {
    if (!req.is('application/json')) {
      return res.status(415).json({ error: 'invalid_content_type' });
    }

    const { idToken, name, message, company } = req.body || {};
    if (!idToken) {
      return res.status(401).json({ error: 'google_required' });
    }

    let googleUser;
    try {
      googleUser = await verifyGoogleToken(idToken);
    } catch (error) {
      const reason = error?.message || 'google_invalid';
      if (reason === 'google_not_configured') {
        return res.status(503).json({ error: 'google_unavailable' });
      }
      return res.status(401).json({ error: 'google_invalid' });
    }

    const fallbackName =
      (typeof name === 'string' && name.trim()) ||
      googleUser.name ||
      googleUser.email.split('@')[0];

    const result = validateContact({
      name: fallbackName,
      email: googleUser.email,
      message,
      company
    });
    if (!result.ok) {
      return res.status(400).json({ error: result.error });
    }

    if (result.isSpam) {
      return res.status(202).json({ ok: true });
    }

    try {
      await mailer.sendContact(result.data);
      return res.status(200).json({ ok: true });
    } catch (error) {
      console.error('contact_send_failed', error);
      return res.status(500).json({ error: 'send_failed' });
    }
  }
);

app.get('/api/health', (_req, res) => {
  res.json({ ok: true });
});

async function startServer() {
  if (!isProd) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'custom'
    });

    app.use(vite.middlewares);

    app.get('*', async (req, res, next) => {
      try {
        const url = req.originalUrl;
        const templatePath = path.resolve(rootDir, 'src', 'index.html');
        let template = await fs.readFile(templatePath, 'utf-8');
        template = await vite.transformIndexHtml(url, template);
        res.status(200).set({ 'Content-Type': 'text/html' }).end(template);
      } catch (error) {
        vite.ssrFixStacktrace(error);
        next(error);
      }
    });
  } else {
    const distDir = path.resolve(rootDir, 'dist');
    app.use(
      express.static(distDir, {
        index: false,
        setHeaders: (res, filePath) => {
          if (filePath.endsWith('.html')) {
            res.setHeader('Cache-Control', 'no-store');
          } else {
            res.setHeader('Cache-Control', 'public, max-age=31536000, immutable');
          }
        }
      })
    );

    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  app.listen(port, () => {
    console.log(`Infinite Dev server running on http://localhost:${port}`);
  });
}

startServer();
