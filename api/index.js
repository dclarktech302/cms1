import express from 'express';
import cors from 'cors';
import crypto from 'crypto';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import {
    S3Client,
    ListObjectsV2Command,
    GetObjectCommand,
    PutObjectCommand,
    HeadObjectCommand,
} from '@aws-sdk/client-s3';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';
import { Client } from '@microsoft/microsoft-graph-client';
import axios from 'axios';
import 'isomorphic-fetch';
import {
    getAccessToken,
    exchangeCodeForTokens,
    getAuthCodeUrl,
    isOneDriveConnected,
} from './auth-config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static(join(__dirname, '../public')));

// ─── AWS S3 ───────────────────────────────────────────────────────────────────

const s3Client = new S3Client({
    region: process.env.AWS_REGION || 'us-east-1',
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
});

const S3_BUCKET = process.env.S3_BUCKET_NAME;
const S3_FOLDER = process.env.S3_FOLDER_PATH || 'gamingclips';

// ─── Session helpers ──────────────────────────────────────────────────────────

function parseCookies(req) {
    const header = req.headers.cookie || '';
    return Object.fromEntries(
        header.split(';')
            .map(c => c.trim())
            .filter(Boolean)
            .map(c => {
                const idx = c.indexOf('=');
                return [c.slice(0, idx).trim(), c.slice(idx + 1).trim()];
            })
    );
}

function signSession(secret) {
    return crypto.createHmac('sha256', secret).update('cms-session').digest('hex');
}

function isValidSession(req) {
    const secret = process.env.ADMIN_SECRET;
    if (!secret) return false;
    const token = parseCookies(req).cms_session;
    if (!token) return false;
    const expected = signSession(secret);
    try {
        return crypto.timingSafeEqual(Buffer.from(token), Buffer.from(expected));
    } catch {
        return false; // length mismatch
    }
}

function requireSession(req, res, next) {
    if (isValidSession(req)) return next();
    res.status(401).json({ error: 'Unauthorized' });
}

// Session guard on all /api routes except public ones
const PUBLIC_API_PATHS = [
    '/api/health',
    '/api/auth/login',
    '/api/auth/status',
    '/api/auth/logout',
    '/api/oauth/authorize-url',
    '/api/oauth/callback',
];

app.use((req, res, next) => {
    if (!req.path.startsWith('/api/')) return next();
    if (PUBLIC_API_PATHS.includes(req.path)) return next();
    if (isValidSession(req)) return next();
    res.status(401).json({ error: 'Unauthorized' });
});

// ─── OneDrive helpers ─────────────────────────────────────────────────────────

const ONEDRIVE_REDIRECT_URI = process.env.ONEDRIVE_REDIRECT_URI;

async function getGraphClient() {
    const token = await getAccessToken();
    return Client.init({
        authProvider: (done) => done(null, token),
    });
}

async function uploadToS3(stream, filename, contentType, contentLength) {
    const key = S3_FOLDER ? `${S3_FOLDER}/${filename}` : filename;
    const params = {
        Bucket: S3_BUCKET,
        Key: key,
        Body: stream,
        ContentType: contentType,
    };
    if (contentLength) params.ContentLength = contentLength;
    await s3Client.send(new PutObjectCommand(params));
    console.log(`✅ Uploaded to S3: ${key}`);
    return key;
}

// ─── Health ───────────────────────────────────────────────────────────────────

app.get('/api/health', (_req, res) => {
    res.json({
        status: 'ok',
        bucket: S3_BUCKET || 'NOT SET',
        folder: S3_FOLDER,
        timestamp: new Date().toISOString(),
    });
});

// ─── Auth ─────────────────────────────────────────────────────────────────────

app.get('/api/auth/status', (req, res) => {
    res.json({ authenticated: isValidSession(req) });
});

app.post('/api/auth/login', (req, res) => {
    const { password } = req.body || {};
    if (!password) return res.status(400).json({ error: 'Password required' });

    const secret = process.env.ADMIN_SECRET;
    if (!secret) return res.status(503).json({ error: 'ADMIN_SECRET not configured' });

    if (password !== secret) return res.status(401).json({ error: 'Incorrect password' });

    const token = signSession(secret);
    const maxAge = 30 * 24 * 60 * 60;
    const secure = process.env.NODE_ENV === 'production' ? '; Secure' : '';
    res.setHeader('Set-Cookie', `cms_session=${token}; HttpOnly; SameSite=Strict; Max-Age=${maxAge}; Path=/${secure}`);
    res.json({ success: true });
});

app.post('/api/auth/logout', (_req, res) => {
    res.setHeader('Set-Cookie', 'cms_session=; HttpOnly; SameSite=Strict; Max-Age=0; Path=/');
    res.json({ success: true });
});

// ─── OAuth (OneDrive) ─────────────────────────────────────────────────────────

app.get('/api/oauth/authorize-url', (_req, res) => {
    try {
        const authUrl = getAuthCodeUrl(ONEDRIVE_REDIRECT_URI);
        res.json({ authUrl });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/oauth/callback', async (req, res) => {
    try {
        const { code, error, error_description } = req.query;

        if (error) {
            return res.status(400).send(errorPage(error_description || error));
        }
        if (!code) return res.status(400).send(errorPage('No authorization code received'));

        await exchangeCodeForTokens(code, ONEDRIVE_REDIRECT_URI);

        res.send(`<!DOCTYPE html>
<html>
<head>
  <title>OneDrive Connected</title>
  <meta http-equiv="refresh" content="3;url=/">
  <style>
    body { font-family: 'IBM Plex Mono', monospace; background: #0f1117; color: #3fb950; max-width: 480px; margin: 80px auto; padding: 24px; text-align: center; }
    h1 { font-size: 2rem; margin-bottom: 12px; }
    p { color: #8b949e; }
    a { color: #d4a574; }
  </style>
</head>
<body>
  <h1>&#10003; OneDrive Connected</h1>
  <p>Tokens saved. Redirecting to dashboard&hellip;</p>
  <p><a href="/">Go now</a></p>
</body>
</html>`);
    } catch (err) {
        res.status(500).send(errorPage(err.message));
    }
});

function errorPage(message) {
    return `<!DOCTYPE html>
<html>
<head>
  <title>Auth Error</title>
  <style>
    body { font-family: monospace; background: #0f1117; color: #f85149; max-width: 480px; margin: 80px auto; padding: 24px; text-align: center; }
    a { color: #d4a574; }
  </style>
</head>
<body>
  <h1>Auth Error</h1>
  <p>${message}</p>
  <p><a href="/">Back to dashboard</a></p>
</body>
</html>`;
}

// ─── OneDrive connection status ───────────────────────────────────────────────

app.get('/api/onedrive/status', (_req, res) => {
    res.json({ connected: isOneDriveConnected() });
});

// ─── OneDrive file browser ────────────────────────────────────────────────────

app.get('/api/onedrive/files', async (_req, res) => {
    try {
        const client = await getGraphClient();

        let files = [];
        let nextUrl = '/me/drive/root:/Videos/Xbox Game DVR:/children' +
            '?$orderby=lastModifiedDateTime desc' +
            '&$top=200';

        while (nextUrl) {
            const response = await client.api(nextUrl).get();
            files = files.concat(response.value || []);
            nextUrl = response['@odata.nextLink'] || null;
        }

        const videoFiles = files.filter(f => f.file?.mimeType?.startsWith('video/'));

        res.json(videoFiles.map(f => ({
            id: f.id,
            name: f.name,
            size: f.size,
            lastModified: f.lastModifiedDateTime,
            mimeType: f.file?.mimeType,
            duration: f.video?.duration,
            width: f.video?.width,
            height: f.video?.height,
            downloadUrl: f['@microsoft.graph.downloadUrl'],
        })));
    } catch (err) {
        if (err.message?.includes('not connected') || err.message?.includes('Authenticate')) {
            return res.status(401).json({ error: 'OneDrive not connected', needsAuth: true });
        }
        if (err.message?.includes('not found') || err.statusCode === 404) {
            return res.status(404).json({ error: 'Xbox Game DVR folder not found in OneDrive' });
        }
        res.status(500).json({ error: err.message });
    }
});

// ─── OneDrive → S3 import ─────────────────────────────────────────────────────

app.post('/api/onedrive/import', async (req, res) => {
    try {
        const { fileId, filename } = req.body || {};
        if (!fileId) return res.status(400).json({ error: 'fileId required' });
        if (!S3_BUCKET) return res.status(500).json({ error: 'S3_BUCKET_NAME not configured' });

        // Get a fresh download URL for the item
        const client = await getGraphClient();
        const item = await client.api(`/me/drive/items/${fileId}`).get();
        const downloadUrl = item['@microsoft.graph.downloadUrl'];
        if (!downloadUrl) return res.status(400).json({ error: 'No download URL available' });

        const name = filename || item.name;

        // Stream from OneDrive directly to S3
        const streamRes = await axios.get(downloadUrl, { responseType: 'stream' });
        const contentType = streamRes.headers['content-type'] || 'video/mp4';
        const contentLength = streamRes.headers['content-length']
            ? parseInt(streamRes.headers['content-length'])
            : undefined;

        const key = await uploadToS3(streamRes.data, name, contentType, contentLength);

        res.json({ success: true, key, filename: name });
    } catch (err) {
        console.error('Import error:', err.message);
        res.status(500).json({ error: err.message });
    }
});

// ─── Webhook (kept for potential future use) ──────────────────────────────────

const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET;

app.post('/api/webhook/onedrive', async (req, res) => {
    try {
        const { validationToken } = req.query;
        if (validationToken) {
            return res.status(200).send(validationToken);
        }

        const notifications = req.body?.value;
        if (!notifications?.length) return res.status(200).json({ message: 'No notifications' });

        res.status(202).json({ message: 'Accepted' });

        for (const notification of notifications) {
            if (notification.clientState !== WEBHOOK_SECRET) continue;
            const itemId = notification.resourceData?.id;
            if (!itemId) continue;

            try {
                const client = await getGraphClient();
                const item = await client.api(`/me/drive/items/${itemId}`).get();
                const downloadUrl = item['@microsoft.graph.downloadUrl'];
                const streamRes = await axios.get(downloadUrl, { responseType: 'stream' });
                const contentType = streamRes.headers['content-type'] || 'video/mp4';
                await uploadToS3(streamRes.data, item.name, contentType);
            } catch (err) {
                console.error(`Error processing ${itemId}:`, err.message);
            }
        }
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Videos (S3 library) ─────────────────────────────────────────────────────

app.get('/api/videos', async (_req, res) => {
    try {
        if (!S3_BUCKET) return res.status(500).json({ error: 'S3_BUCKET_NAME not configured' });

        const response = await s3Client.send(
            new ListObjectsV2Command({
                Bucket: S3_BUCKET,
                Prefix: S3_FOLDER ? `${S3_FOLDER}/` : '',
            })
        );

        if (!response.Contents) return res.json([]);

        const videos = response.Contents.filter(item => !item.Key.endsWith('/'))
            .map(item => ({
                key: item.Key,
                filename: item.Key.split('/').pop(),
                size_bytes: item.Size,
                last_modified: item.LastModified,
                url: `/api/video-url?key=${encodeURIComponent(item.Key)}`,
            }))
            .sort((a, b) => new Date(b.last_modified) - new Date(a.last_modified));

        res.json(videos);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/video-url', async (req, res) => {
    try {
        const { key } = req.query;
        if (!key) return res.status(400).json({ error: 'key parameter required' });
        if (!S3_BUCKET) return res.status(500).json({ error: 'S3_BUCKET_NAME not configured' });
        if (S3_FOLDER && !key.startsWith(`${S3_FOLDER}/`))
            return res.status(403).json({ error: 'Access denied' });

        const url = await getSignedUrl(
            s3Client,
            new GetObjectCommand({ Bucket: S3_BUCKET, Key: key }),
            { expiresIn: 3600 }
        );
        res.json({ url });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/test-video/:key', async (req, res) => {
    try {
        const fullKey = S3_FOLDER ? `${S3_FOLDER}/${req.params.key}` : req.params.key;
        const metadata = await s3Client.send(
            new HeadObjectCommand({ Bucket: S3_BUCKET, Key: fullKey })
        );
        const url = await getSignedUrl(
            s3Client,
            new GetObjectCommand({ Bucket: S3_BUCKET, Key: fullKey }),
            { expiresIn: 3600 }
        );
        res.json({ key: fullKey, url, metadata });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Login page ───────────────────────────────────────────────────────────────

app.get('/login', (_req, res) => {
    res.sendFile(join(__dirname, '../public/login.html'));
});

// ─── Catch-all ────────────────────────────────────────────────────────────────

app.get('*', (_req, res) => {
    res.sendFile(join(__dirname, '../public/index.html'));
});

// ─── Local dev server ─────────────────────────────────────────────────────────

if (process.env.NODE_ENV !== 'production') {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => console.log(`🚀 Server running at http://localhost:${PORT}`));
}

export default app;
