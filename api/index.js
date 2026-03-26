import express from 'express';
import cors from 'cors';
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
} from './auth-config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();

app.use(cors());
app.use(express.json());

// Serve static files from /public with absolute path
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

// ─── OneDrive config ──────────────────────────────────────────────────────────

const ONEDRIVE_REDIRECT_URI = process.env.ONEDRIVE_REDIRECT_URI;
const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET;

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function getGraphClient() {
    const token = await getAccessToken();
    return Client.init({
        authProvider: (done) => done(null, token),
    });
}

async function downloadFromOneDrive(itemId) {
    const client = await getGraphClient();
    const item = await client.api(`/me/drive/items/${itemId}`).get();
    const response = await axios.get(item['@microsoft.graph.downloadUrl'], {
        responseType: 'arraybuffer',
    });
    return {
        buffer: response.data,
        filename: item.name,
        contentType: response.headers['content-type'],
    };
}

async function uploadToS3(buffer, filename, contentType) {
    const key = S3_FOLDER ? `${S3_FOLDER}/${filename}` : filename;
    await s3Client.send(
        new PutObjectCommand({
            Bucket: S3_BUCKET,
            Key: key,
            Body: buffer,
            ContentType: contentType,
        })
    );
    console.log(`✅ Uploaded to S3: ${key}`);
    return key;
}

async function deleteFromOneDrive(itemId) {
    const client = await getGraphClient();
    await client.api(`/me/drive/items/${itemId}`).delete();
    console.log(`🗑️  Deleted from OneDrive: ${itemId}`);
}

// ─── Health ───────────────────────────────────────────────────────────────────

app.get('/api/health', (req, res) => {
    res.json({
        status: 'ok',
        bucket: S3_BUCKET || 'NOT SET',
        folder: S3_FOLDER,
        timestamp: new Date().toISOString(),
    });
});

// ─── OAuth ────────────────────────────────────────────────────────────────────

app.get('/api/oauth/authorize-url', async (_req, res) => {
    try {
        if (!ONEDRIVE_REDIRECT_URI) {
            return res.status(500).json({
                error: 'ONEDRIVE_REDIRECT_URI is not set in environment variables',
            });
        }
        const authUrl = getAuthCodeUrl(ONEDRIVE_REDIRECT_URI);
        if (!authUrl || typeof authUrl !== 'string') {
            return res.status(500).json({ error: 'Failed to build authorization URL' });
        }
        res.json({ authUrl });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/oauth/callback', async (req, res) => {
    try {
        const { code } = req.query;
        if (!code) return res.status(400).json({ error: 'No authorization code' });

        const tokens = await exchangeCodeForTokens(code, ONEDRIVE_REDIRECT_URI);

        res.send(`<!DOCTYPE html>
<html>
<head>
  <title>Authorization Successful</title>
  <style>
    body { font-family: monospace; background: #0d0d0d; color: #f0a500; max-width: 640px; margin: 60px auto; padding: 24px; }
    h1 { color: #f0a500; } code { background: #1a1a1a; padding: 4px 8px; border-radius: 3px; display: block; margin: 8px 0; word-break: break-all; color: #ccc; }
    a { color: #f0a500; }
  </style>
</head>
<body>
  <h1>✓ OneDrive Connected</h1>
  <p>Copy these tokens into your Vercel environment variables:</p>
  <p><strong>ONEDRIVE_ACCESS_TOKEN:</strong></p>
  <code>${tokens.accessToken.substring(0, 60)}...</code>
  <p><strong>ONEDRIVE_REFRESH_TOKEN:</strong></p>
  <code>${tokens.refreshToken ? tokens.refreshToken.substring(0, 60) + '...' : 'Not available'}</code>
  <p>Expires: ${tokens.expiresOn ? new Date(tokens.expiresOn).toLocaleString() : 'Unknown'}</p>
  <p><a href="/">← Back to CMS</a></p>
</body>
</html>`);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Webhook ──────────────────────────────────────────────────────────────────

app.post('/api/webhook/onedrive', async (req, res) => {
    try {
        const { validationToken } = req.query;
        if (validationToken) {
            console.log('🔗 Webhook validation OK');
            return res.status(200).send(validationToken);
        }

        const notifications = req.body?.value;
        if (!notifications?.length) return res.status(200).json({ message: 'No notifications' });

        console.log(`📬 ${notifications.length} notification(s) received`);

        // Respond to Microsoft immediately (must be < 3s)
        res.status(202).json({ message: 'Accepted' });

        // Process async after response
        for (const notification of notifications) {
            if (notification.clientState !== WEBHOOK_SECRET) {
                console.warn('⚠️  Invalid clientState — skipping');
                continue;
            }

            const itemId = notification.resourceData?.id;
            if (!itemId) continue;

            try {
                console.log(`⬇️  Downloading from OneDrive: ${itemId}`);
                const { buffer, filename, contentType } = await downloadFromOneDrive(itemId);

                console.log(`⬆️  Uploading to S3: ${filename}`);
                await uploadToS3(buffer, filename, contentType);

                // Optional: delete from OneDrive after upload
                // await deleteFromOneDrive(itemId);

                console.log(`✅ Synced: ${filename}`);
            } catch (err) {
                console.error(`❌ Error processing ${itemId}:`, err.message);
            }
        }
    } catch (err) {
        console.error('Webhook error:', err.message);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/webhook/setup-onedrive', async (req, res) => {
    try {
        const client = await getGraphClient();

        let folderId;
        try {
            const folder = await client.api('/me/drive/root:/Videos/Xbox Game DVR').get();
            folderId = folder.id;
        } catch {
            return res.status(404).json({
                error: 'Xbox Game DVR folder not found',
                message: 'Ensure "Videos/Xbox Game DVR" exists in your OneDrive',
            });
        }

        const expirationDateTime = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString();
        const notificationUrl =
            req.body.notificationUrl ||
            `${process.env.VERCEL_URL ? 'https://' + process.env.VERCEL_URL : 'http://localhost:3000'}/api/webhook/onedrive`;

        const subscription = {
            changeType: 'created,updated',
            notificationUrl,
            resource: `/me/drive/items/${folderId}`,
            expirationDateTime,
            clientState: WEBHOOK_SECRET,
        };

        const result = await client.api('/subscriptions').post(subscription);

        res.json({
            success: true,
            subscription: {
                id: result.id,
                resource: result.resource,
                expirationDateTime: result.expirationDateTime,
                notificationUrl: result.notificationUrl,
            },
            note: 'Subscription expires in 24h — renew periodically.',
        });
    } catch (err) {
        if (err.message?.includes('No valid token')) {
            return res.status(401).json({
                error: 'Not authenticated',
                message: 'Complete OAuth flow first via /api/oauth/authorize-url',
            });
        }
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/webhook/list-subscriptions', async (req, res) => {
    try {
        const client = await getGraphClient();
        const result = await client.api('/subscriptions').get();
        res.json({ count: result.value.length, subscriptions: result.value });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.delete('/api/webhook/subscription/:id', async (req, res) => {
    try {
        const client = await getGraphClient();
        await client.api(`/subscriptions/${req.params.id}`).delete();
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// ─── Videos ───────────────────────────────────────────────────────────────────

app.get('/api/videos', async (req, res) => {
    try {
        if (!S3_BUCKET) return res.status(500).json({ error: 'S3_BUCKET_NAME not configured' });

        const response = await s3Client.send(
            new ListObjectsV2Command({
                Bucket: S3_BUCKET,
                Prefix: S3_FOLDER ? `${S3_FOLDER}/` : '',
            })
        );

        if (!response.Contents) return res.json([]);

        const videos = response.Contents.filter((item) => !item.Key.endsWith('/'))
            .map((item) => ({
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
        if (S3_FOLDER && !key.startsWith(`${S3_FOLDER}/`))
            return res.status(403).json({ error: 'Access denied' });

        const CLOUDFRONT_DOMAIN = process.env.CLOUDFRONT_DOMAIN;

        if (CLOUDFRONT_DOMAIN) {
            // Serve via CloudFront — supports range requests natively, no expiry
            const url = `https://${CLOUDFRONT_DOMAIN}/${key}`;
            return res.json({ url });
        }

        // Fallback: presigned S3 URL (no CloudFront configured)
        if (!S3_BUCKET) return res.status(500).json({ error: 'S3_BUCKET_NAME not configured' });
        const url = await getSignedUrl(
            s3Client,
            new GetObjectCommand({
                Bucket: S3_BUCKET,
                Key: key,
                // No checksum mode — prevents 416 on range requests
            }),
            { expiresIn: 3600, unhoistableHeaders: new Set(['x-amz-checksum-mode']) }
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
        const CLOUDFRONT_DOMAIN = process.env.CLOUDFRONT_DOMAIN;
        const url = CLOUDFRONT_DOMAIN
            ? `https://${CLOUDFRONT_DOMAIN}/${fullKey}`
            : await getSignedUrl(
                s3Client,
                new GetObjectCommand({ Bucket: S3_BUCKET, Key: fullKey }),
                { expiresIn: 3600, unhoistableHeaders: new Set(['x-amz-checksum-mode']) }
            );
        res.json({ key: fullKey, url, metadata });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});


// ─── OneDrive Browser ─────────────────────────────────────────────────────────

// List files in Xbox Game DVR folder
app.get('/api/onedrive/files', async (req, res) => {
    try {
        const client = await getGraphClient();

        // Get the Xbox Game DVR folder
        let folder;
        try {
            folder = await client.api('/me/drive/root:/Videos/Xbox Game DVR').get();
        } catch {
            return res.status(404).json({
                error: 'Xbox Game DVR folder not found',
                message: 'Ensure "Videos/Xbox Game DVR" exists in your OneDrive',
            });
        }

        // List children, filter to video files only
        const children = await client
            .api(`/me/drive/items/${folder.id}/children`)
            .select('id,name,size,lastModifiedDateTime,file,video')
            .orderby('lastModifiedDateTime desc')
            .top(100)
            .get();

        const videoExtensions = /\.(mp4|mov|avi|mkv|wmv|m4v|webm)$/i;

        const files = (children.value || [])
            .filter(item => item.file && videoExtensions.test(item.name))
            .map(item => ({
                id: item.id,
                name: item.name,
                size_bytes: item.size,
                last_modified: item.lastModifiedDateTime,
                duration_ms: item.video?.duration || null,
            }));

        res.json({ count: files.length, files });
    } catch (err) {
        if (err.message?.includes('No valid token')) {
            return res.status(401).json({ error: 'Not authenticated' });
        }
        res.status(500).json({ error: err.message });
    }
});

// Import a single file from OneDrive → S3 (keeps OneDrive copy)
app.post('/api/onedrive/import', async (req, res) => {
    try {
        const { itemId, filename } = req.body;
        if (!itemId || !filename) {
            return res.status(400).json({ error: 'itemId and filename are required' });
        }
        if (!S3_BUCKET) {
            return res.status(500).json({ error: 'S3_BUCKET_NAME not configured' });
        }

        // Check if already exists in S3
        const s3Key = S3_FOLDER ? `${S3_FOLDER}/${filename}` : filename;
        try {
            await s3Client.send(new HeadObjectCommand({ Bucket: S3_BUCKET, Key: s3Key }));
            return res.status(409).json({
                error: 'Already imported',
                message: `${filename} already exists in S3`,
                s3Key,
            });
        } catch (headErr) {
            // 404 = not found = good, proceed with import
            if (headErr.name !== 'NotFound' && headErr.$metadata?.httpStatusCode !== 404) {
                throw headErr;
            }
        }

        console.log(`Importing from OneDrive: ${filename} (${itemId})`);

        // Download from OneDrive
        const { buffer, contentType } = await downloadFromOneDrive(itemId);

        // Upload to S3
        await s3Client.send(
            new PutObjectCommand({
                Bucket: S3_BUCKET,
                Key: s3Key,
                Body: buffer,
                ContentType: contentType || 'video/mp4',
            })
        );

        console.log(`Imported to S3: ${s3Key}`);

        res.json({
            success: true,
            filename,
            s3Key,
            message: `${filename} imported successfully`,
        });
    } catch (err) {
        console.error('Import error:', err.message);
        if (err.message?.includes('No valid token')) {
            return res.status(401).json({ error: 'Not authenticated' });
        }
        res.status(500).json({ error: err.message });
    }
});

// ─── Catch-all → index.html ───────────────────────────────────────────────────

app.get('*', (req, res) => {
    res.sendFile(join(__dirname, '../public/index.html'));
});

// ─── Start (local dev only — Vercel uses export) ──────────────────────────────

if (process.env.NODE_ENV !== 'production') {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
}

export default app;