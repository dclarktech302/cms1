import express from 'express';
import cors from 'cors';
import crypto from 'crypto';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import {
    S3Client,
    CopyObjectCommand,
    DeleteObjectCommand,
    GetObjectCommand,
    PutObjectCommand,
    HeadBucketCommand,
} from '@aws-sdk/client-s3';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';
import axios from 'axios';
import supabase from './supabase-client.js';
import { getAuthCodeUrl, exchangeCodeForTokens } from './auth-config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static(join(__dirname, '../public')));

// ─── AWS S3 ───────────────────────────────────────────────────────────────────

const s3 = new S3Client({
    region: process.env.AWS_REGION || 'us-east-1',
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
});

const S3_BUCKET = process.env.S3_BUCKET_NAME || process.env.AWS_S3_BUCKET_NAME;

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
    try { return crypto.timingSafeEqual(Buffer.from(token), Buffer.from(expected)); }
    catch { return false; }
}

// Session guard — all /api/* except the allowlist below
const PUBLIC_API_PATHS = [
    '/api/health',
    '/api/auth/login',
    '/api/auth/status',
    '/api/auth/logout',
    '/api/oauth/authorize-url',
    '/api/oauth/callback',
    '/api/clips/ingest',  // protected by CMS_INGEST_SECRET instead
];

app.use((req, res, next) => {
    if (!req.path.startsWith('/api/')) return next();
    if (PUBLIC_API_PATHS.includes(req.path)) return next();
    if (isValidSession(req)) return next();
    res.status(401).json({ error: 'Unauthorized' });
});

// ─── Health ───────────────────────────────────────────────────────────────────

app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok', bucket: S3_BUCKET || 'NOT SET', timestamp: new Date().toISOString() });
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
    const secure = process.env.NODE_ENV === 'production' ? '; Secure' : '';
    res.setHeader('Set-Cookie', `cms_session=${token}; HttpOnly; SameSite=Strict; Max-Age=${30 * 24 * 3600}; Path=/${secure}`);
    res.json({ success: true });
});

app.post('/api/auth/logout', (_req, res) => {
    res.setHeader('Set-Cookie', 'cms_session=; HttpOnly; SameSite=Strict; Max-Age=0; Path=/');
    res.json({ success: true });
});

// ─── OneDrive OAuth (initial token only) ─────────────────────────────────────

function getRedirectUri(req) {
    const configured = process.env.ONEDRIVE_REDIRECT_URI;
    if (configured && !configured.includes('localhost')) return configured;
    if (process.env.VERCEL_PROJECT_PRODUCTION_URL) {
        return `https://${process.env.VERCEL_PROJECT_PRODUCTION_URL}/api/oauth/callback`;
    }
    const proto = req.headers['x-forwarded-proto'] || req.protocol || 'https';
    return `${proto}://${req.headers.host}/api/oauth/callback`;
}

app.get('/api/oauth/authorize-url', (req, res) => {
    try {
        res.json({ authUrl: getAuthCodeUrl(getRedirectUri(req)) });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/oauth/callback', async (req, res) => {
    try {
        const { code, state, error, error_description } = req.query;
        if (error) return res.status(400).send(htmlPage('Auth Error', `<p style="color:#f85149">${error_description || error}</p><p><a href="/">Back</a></p>`));
        if (!code) return res.status(400).send(htmlPage('Auth Error', '<p>No authorization code received</p>'));

        let redirectUri = getRedirectUri(req);
        if (state) { try { redirectUri = Buffer.from(state, 'base64').toString(); } catch {} }

        const tokens = await exchangeCodeForTokens(code, redirectUri);

        res.send(htmlPage('OneDrive Connected', `
            <h2 style="color:#3fb950">&#10003; OneDrive Connected</h2>
            <p>Copy the refresh token below into GitHub Secrets as <code>ONEDRIVE_REFRESH_TOKEN</code>:</p>
            <textarea style="width:100%;height:80px;background:#0d1117;color:#ccc;border:1px solid #30363d;padding:8px;font-family:monospace;font-size:11px;border-radius:4px"
                onclick="this.select()">${tokens.refresh_token || 'Not returned — re-authenticate with offline_access scope'}</textarea>
            <p style="color:#8b949e;font-size:0.85rem">After adding to GitHub Secrets, the Action will auto-rotate this token going forward.</p>
            <p><a href="/">Back to dashboard</a></p>
        `));
    } catch (err) {
        res.status(500).send(htmlPage('Auth Error', `<p style="color:#f85149">${err.message}</p>`));
    }
});

function htmlPage(title, body) {
    return `<!DOCTYPE html><html><head><title>${title}</title>
<style>body{font-family:monospace;background:#0f1117;color:#e6edf3;max-width:600px;margin:60px auto;padding:24px}a{color:#d4a574}code{background:#161b22;padding:2px 6px;border-radius:3px}</style>
</head><body>${body}</body></html>`;
}

// ─── Clips — ingest (called by GitHub Action) ────────────────────────────────

app.post('/api/clips/ingest', async (req, res) => {
    const secret = req.headers['x-ingest-secret'];
    if (!secret || secret !== process.env.CMS_INGEST_SECRET) {
        return res.status(401).json({ error: 'Invalid ingest secret' });
    }

    const { filename, s3Key, size, onedrive_id, recorded_at } = req.body || {};
    if (!filename || !s3Key) return res.status(400).json({ error: 'filename and s3Key required' });

    const { data, error } = await supabase
        .from('clips')
        .insert({
            filename,
            s3_key: s3Key,
            size_bytes: size || null,
            recorded_at: recorded_at || null,
        })
        .select('id, s3_key')
        .single();

    if (error) {
        if (error.code === '23505') return res.status(409).json({ error: 'Clip already ingested' });
        console.error('Supabase insert error:', error);
        return res.status(500).json({ error: error.message });
    }

    res.status(201).json({ id: data.id, s3_key: data.s3_key });
});

// ─── Clips — list ─────────────────────────────────────────────────────────────

app.get('/api/clips', async (req, res) => {
    try {
        const { status } = req.query;
        let query = supabase.from('clips').select('*').order('created_at', { ascending: false });
        if (status) {
            const statuses = status.split(',');
            query = statuses.length === 1
                ? query.eq('status', statuses[0])
                : query.in('status', statuses);
        }
        const { data, error } = await query;
        if (error) throw error;
        res.json(data);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Clips — presigned preview URL ───────────────────────────────────────────

app.get('/api/clips/:id/preview', async (req, res) => {
    try {
        const { data: clip, error } = await supabase
            .from('clips').select('s3_key, trimmed_s3_key').eq('id', req.params.id).single();
        if (error || !clip) return res.status(404).json({ error: 'Clip not found' });

        const key = clip.trimmed_s3_key || clip.s3_key;
        const url = await getSignedUrl(s3, new GetObjectCommand({ Bucket: S3_BUCKET, Key: key }), { expiresIn: 1800 });
        res.json({ url, expiresIn: 1800 });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Clips — presigned upload URL (for trimmed output) ───────────────────────

app.get('/api/clips/:id/presigned-upload', async (req, res) => {
    try {
        const { data: clip, error } = await supabase
            .from('clips').select('filename').eq('id', req.params.id).single();
        if (error || !clip) return res.status(404).json({ error: 'Clip not found' });

        const key = `trimmed/${clip.filename}`;
        const url = await getSignedUrl(
            s3,
            new PutObjectCommand({ Bucket: S3_BUCKET, Key: key, ContentType: 'video/mp4' }),
            { expiresIn: 3600 }
        );
        res.json({ url, key, expiresIn: 3600 });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Clips — accept ───────────────────────────────────────────────────────────

app.post('/api/clips/:id/accept', async (req, res) => {
    try {
        const { data: clip, error: fetchErr } = await supabase
            .from('clips').select('s3_key, filename').eq('id', req.params.id).single();
        if (fetchErr || !clip) return res.status(404).json({ error: 'Clip not found' });

        const newKey = `library/${clip.filename}`;

        // Copy pending/ → library/
        await s3.send(new CopyObjectCommand({
            Bucket: S3_BUCKET,
            CopySource: `${S3_BUCKET}/${clip.s3_key}`,
            Key: newKey,
        }));

        // Delete from pending/
        await s3.send(new DeleteObjectCommand({ Bucket: S3_BUCKET, Key: clip.s3_key }));

        // Update Supabase
        const { error } = await supabase
            .from('clips')
            .update({ status: 'accepted', s3_key: newKey })
            .eq('id', req.params.id);
        if (error) throw error;

        res.json({ success: true, s3_key: newKey });
    } catch (err) {
        console.error('Accept error:', err.message);
        res.status(500).json({ error: err.message });
    }
});

// ─── Clips — discard ─────────────────────────────────────────────────────────

app.delete('/api/clips/:id', async (req, res) => {
    try {
        const { data: clip, error: fetchErr } = await supabase
            .from('clips').select('s3_key, trimmed_s3_key').eq('id', req.params.id).single();
        if (fetchErr || !clip) return res.status(404).json({ error: 'Clip not found' });

        // Delete from S3 (both keys if trimmed version exists)
        const deleteOps = [
            s3.send(new DeleteObjectCommand({ Bucket: S3_BUCKET, Key: clip.s3_key })),
        ];
        if (clip.trimmed_s3_key) {
            deleteOps.push(s3.send(new DeleteObjectCommand({ Bucket: S3_BUCKET, Key: clip.trimmed_s3_key })));
        }
        await Promise.allSettled(deleteOps);

        // Mark discarded in Supabase
        const { error } = await supabase
            .from('clips')
            .update({ status: 'discarded' })
            .eq('id', req.params.id);
        if (error) throw error;

        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Clips — patch (trim complete, status update) ────────────────────────────

app.patch('/api/clips/:id', async (req, res) => {
    try {
        const allowed = ['status', 'trimmed_s3_key', 'tiktok_draft_id'];
        const updates = Object.fromEntries(
            Object.entries(req.body || {}).filter(([k]) => allowed.includes(k))
        );
        if (!Object.keys(updates).length) return res.status(400).json({ error: 'No valid fields' });

        const { error } = await supabase.from('clips').update(updates).eq('id', req.params.id);
        if (error) throw error;
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Sync — trigger GitHub Action ────────────────────────────────────────────

app.post('/api/sync/trigger', async (_req, res) => {
    try {
        const token = process.env.GITHUB_PAT;
        if (!token) return res.status(503).json({ error: 'GITHUB_PAT not configured' });

        await axios.post(
            'https://api.github.com/repos/dclarktech302/cms1/actions/workflows/sync-clips.yml/dispatches',
            { ref: 'main' },
            {
                headers: {
                    Authorization: `Bearer ${token}`,
                    Accept: 'application/vnd.github+json',
                },
            }
        );
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.response?.data?.message || err.message });
    }
});

// ─── Status — connectivity checks ────────────────────────────────────────────

app.get('/api/status', async (_req, res) => {
    const results = {};

    // Supabase: check by querying clips count
    try {
        const { error } = await supabase.from('clips').select('id', { count: 'exact', head: true });
        results.supabase = error ? { ok: false, error: error.message } : { ok: true };
    } catch (err) {
        results.supabase = { ok: false, error: err.message };
    }

    // S3
    try {
        await s3.send(new HeadBucketCommand({ Bucket: S3_BUCKET }));
        results.s3 = { ok: true, bucket: S3_BUCKET };
    } catch (err) {
        results.s3 = { ok: false, error: err.message, bucket: S3_BUCKET };
    }

    // Last sync: most recent clip created_at
    try {
        const { data } = await supabase
            .from('clips').select('created_at').order('created_at', { ascending: false }).limit(1);
        results.lastSync = data?.[0]?.created_at || null;
    } catch {
        results.lastSync = null;
    }

    results.githubPat = !!process.env.GITHUB_PAT;

    res.json(results);
});

// ─── Login page ───────────────────────────────────────────────────────────────

app.get('/login', (_req, res) => res.sendFile(join(__dirname, '../public/login.html')));

// ─── Catch-all ────────────────────────────────────────────────────────────────

app.get('*', (_req, res) => res.sendFile(join(__dirname, '../public/index.html')));

// ─── Local dev ────────────────────────────────────────────────────────────────

if (process.env.NODE_ENV !== 'production') {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => console.log(`🚀 http://localhost:${PORT}`));
}

export default app;
