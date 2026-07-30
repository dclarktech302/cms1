#!/usr/bin/env node
// Runs in GitHub Actions every 15 minutes.
// Authenticates with OneDrive, lists new Xbox clips, streams them to S3,
// and notifies the CMS ingest endpoint. Self-rotates the refresh token.

import axios from 'axios';
import { S3Client, PutObjectCommand } from '@aws-sdk/client-s3';

const TENANT_ID = process.env.AZURE_TENANT_ID || 'common';
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const SCOPES = 'https://graph.microsoft.com/Files.ReadWrite offline_access';
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const XBOX_FOLDER = '/me/drive/root:/Videos/Xbox Game DVR:/children';

// OIDC credentials are injected automatically by aws-actions/configure-aws-credentials
const s3 = new S3Client({ region: process.env.AWS_REGION || 'us-east-1' });

// ─── Token refresh ────────────────────────────────────────────────────────────

async function refreshAccessToken() {
    const res = await axios.post(
        TOKEN_URL,
        new URLSearchParams({
            client_id: process.env.AZURE_CLIENT_ID,
            client_secret: process.env.AZURE_CLIENT_SECRET,
            refresh_token: process.env.ONEDRIVE_REFRESH_TOKEN,
            grant_type: 'refresh_token',
            scope: SCOPES,
        }).toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token: newRefresh } = res.data;

    // Self-rotate: update GitHub Secret when Microsoft issues a new refresh token
    if (newRefresh && newRefresh !== process.env.ONEDRIVE_REFRESH_TOKEN) {
        await rotateGitHubSecret('ONEDRIVE_REFRESH_TOKEN', newRefresh);
        process.env.ONEDRIVE_REFRESH_TOKEN = newRefresh;
    }

    return access_token;
}

// ─── GitHub Secrets rotation ──────────────────────────────────────────────────

async function rotateGitHubSecret(name, value) {
    const token = process.env.GH_TOKEN;
    const repo = process.env.GH_REPO;
    if (!token || !repo) {
        console.warn(`⚠️  Cannot rotate ${name}: GH_TOKEN or GH_REPO not set`);
        return;
    }

    try {
        // Get repo public key for encryption
        const keyRes = await axios.get(
            `https://api.github.com/repos/${repo}/actions/secrets/public-key`,
            { headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json' } }
        );
        const { key: publicKey, key_id } = keyRes.data;
        const encrypted = await encryptSecret(publicKey, value);

        await axios.put(
            `https://api.github.com/repos/${repo}/actions/secrets/${name}`,
            { encrypted_value: encrypted, key_id },
            { headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json' } }
        );
        console.log(`🔑 Rotated GitHub Secret: ${name}`);
    } catch (err) {
        console.error(`⚠️  Failed to rotate ${name}: ${err.response?.data?.message || err.message}`);
    }
}

async function encryptSecret(publicKey, secretValue) {
    const sodium = await import('libsodium-wrappers');
    await sodium.default.ready;
    const lib = sodium.default;
    const keyBytes = Buffer.from(publicKey, 'base64');
    const messageBytes = Buffer.from(secretValue);
    const encrypted = lib.crypto_box_seal(messageBytes, keyBytes);
    return Buffer.from(encrypted).toString('base64');
}

// ─── OneDrive file listing ────────────────────────────────────────────────────

async function listRecentFiles(accessToken) {
    // List files modified in the last hour — the ingest endpoint handles 409 deduplication
    const since = new Date(Date.now() - 60 * 60 * 1000).toISOString();

    let files = [];
    let url = `${GRAPH_BASE}${XBOX_FOLDER}?$orderby=lastModifiedDateTime desc&$top=200`;

    while (url) {
        const res = await axios.get(url, {
            headers: { Authorization: `Bearer ${accessToken}` },
        });
        const page = (res.data.value || []).filter(f =>
            f.file?.mimeType?.startsWith('video/') && f.lastModifiedDateTime > since
        );
        files = files.concat(page);

        // Stop paginating once items are older than our window
        const hasMore = page.length === (res.data.value?.length ?? 0);
        url = hasMore ? (res.data['@odata.nextLink'] || null) : null;
    }

    return files;
}

// ─── S3 upload (streaming) ────────────────────────────────────────────────────

async function uploadToS3(downloadUrl, filename) {
    const streamRes = await axios.get(downloadUrl, { responseType: 'stream' });
    const key = `pending/${filename}`;

    await s3.send(new PutObjectCommand({
        Bucket: process.env.AWS_S3_BUCKET_NAME,
        Key: key,
        Body: streamRes.data,
        ContentType: streamRes.headers['content-type'] || 'video/mp4',
        ContentLength: parseInt(streamRes.headers['content-length'] || '0') || undefined,
    }));

    return key;
}

// ─── CMS ingest notification ──────────────────────────────────────────────────

async function notifyIngest(file, s3Key) {
    const url = process.env.CMS_INGEST_URL;
    if (!url) return;

    await axios.post(url, {
        filename: file.name,
        s3Key,
        size: file.size,
        onedrive_id: file.id,
        recorded_at: file.fileSystemInfo?.createdDateTime || file.lastModifiedDateTime,
    }, {
        headers: { 'x-ingest-secret': process.env.CMS_INGEST_SECRET },
        timeout: 15000,
    });
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main() {
    console.log(`🎮 Xbox clip sync — ${new Date().toISOString()}`);

    if (!process.env.ONEDRIVE_REFRESH_TOKEN) {
        console.error('❌ ONEDRIVE_REFRESH_TOKEN not set in GitHub Secrets');
        process.exit(1);
    }

    const accessToken = await refreshAccessToken();
    console.log('✅ OneDrive token refreshed');

    const files = await listRecentFiles(accessToken);
    console.log(`📁 ${files.length} new video file(s) in the last hour`);

    if (!files.length) {
        console.log('✅ Nothing to sync');
        return;
    }

    let synced = 0;
    let skipped = 0;
    let failed = 0;

    for (const file of files) {
        try {
            console.log(`⬇️  ${file.name} (${(file.size / 1024 / 1024).toFixed(1)} MB)`);
            const s3Key = await uploadToS3(file['@microsoft.graph.downloadUrl'], file.name);

            try {
                await notifyIngest(file, s3Key);
                console.log(`✅ Synced: ${file.name}`);
                synced++;
            } catch (err) {
                if (err.response?.status === 409) {
                    console.log(`⏭️  Already in CMS: ${file.name}`);
                    skipped++;
                } else {
                    throw err;
                }
            }
        } catch (err) {
            console.error(`❌ ${file.name}: ${err.message}`);
            failed++;
        }
    }

    console.log(`\n✅ Done — synced: ${synced}, skipped: ${skipped}, failed: ${failed}`);
    if (failed > 0) process.exit(1);
}

main().catch(err => {
    console.error('Fatal:', err.message);
    process.exit(1);
});
