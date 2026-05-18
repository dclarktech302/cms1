import axios from 'axios';

const TENANT_ID = process.env.ONEDRIVE_TENANT_ID || process.env.AZURE_TENANT_ID || 'common';
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const SCOPES = 'https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/Files.ReadWrite.All offline_access';

// ─── Vercel env var persistence ───────────────────────────────────────────────

async function writeVercelEnv(key, value) {
    // Always update in-process so the current request benefits immediately
    process.env[key] = value;

    const projectId = process.env.VERCEL_PROJECT_ID;
    const apiToken = process.env.VERCEL_API_TOKEN;
    if (!projectId || !apiToken) return; // local dev: in-memory only

    try {
        const listRes = await axios.get(
            `https://api.vercel.com/v9/projects/${projectId}/env`,
            { headers: { Authorization: `Bearer ${apiToken}` } }
        );
        const envVar = listRes.data.envs.find(e => e.key === key);

        if (envVar) {
            await axios.patch(
                `https://api.vercel.com/v9/projects/${projectId}/env/${envVar.id}`,
                { value },
                { headers: { Authorization: `Bearer ${apiToken}` } }
            );
        } else {
            await axios.post(
                `https://api.vercel.com/v9/projects/${projectId}/env`,
                { key, value, type: 'encrypted', target: ['production', 'preview', 'development'] },
                { headers: { Authorization: `Bearer ${apiToken}` } }
            );
        }
    } catch (err) {
        console.error(`⚠️  Failed to persist ${key} to Vercel:`, err.message);
    }
}

async function persistTokens(accessToken, refreshToken, expiresIn) {
    const expiresAt = new Date(Date.now() + expiresIn * 1000).toISOString();
    await Promise.all([
        writeVercelEnv('ONEDRIVE_ACCESS_TOKEN', accessToken),
        writeVercelEnv('ONEDRIVE_REFRESH_TOKEN', refreshToken),
        writeVercelEnv('ONEDRIVE_TOKEN_EXPIRES_AT', expiresAt),
    ]);
    console.log('✅ OneDrive tokens persisted');
}

// ─── OAuth helpers ────────────────────────────────────────────────────────────

export function getAuthCodeUrl(redirectUri) {
    const params = new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        response_type: 'code',
        redirect_uri: redirectUri,
        scope: SCOPES,
        response_mode: 'query',
        prompt: 'consent',
    });
    return `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?${params}`;
}

export async function exchangeCodeForTokens(code, redirectUri) {
    const res = await axios.post(
        TOKEN_URL,
        new URLSearchParams({
            client_id: process.env.AZURE_CLIENT_ID,
            client_secret: process.env.AZURE_CLIENT_SECRET,
            code,
            redirect_uri: redirectUri,
            grant_type: 'authorization_code',
            scope: SCOPES,
        }).toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token, expires_in } = res.data;
    await persistTokens(access_token, refresh_token, expires_in);
    return { accessToken: access_token, refreshToken: refresh_token };
}

// ─── Token access (auto-refresh) ─────────────────────────────────────────────

export async function getAccessToken() {
    const accessToken = process.env.ONEDRIVE_ACCESS_TOKEN;
    const refreshToken = process.env.ONEDRIVE_REFRESH_TOKEN;
    const expiresAt = process.env.ONEDRIVE_TOKEN_EXPIRES_AT;

    // Return existing token if still valid with 60s buffer
    if (accessToken && expiresAt) {
        if (new Date(expiresAt) > new Date(Date.now() + 60_000)) {
            return accessToken;
        }
    }

    if (!refreshToken) {
        throw new Error('OneDrive not connected. Authenticate via the dashboard.');
    }

    console.log('🔄 Refreshing OneDrive access token...');
    try {
        const res = await axios.post(
            TOKEN_URL,
            new URLSearchParams({
                client_id: process.env.AZURE_CLIENT_ID,
                client_secret: process.env.AZURE_CLIENT_SECRET,
                refresh_token: refreshToken,
                grant_type: 'refresh_token',
                scope: SCOPES,
            }).toString(),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const { access_token, refresh_token: new_refresh, expires_in } = res.data;
        await persistTokens(access_token, new_refresh, expires_in);
        return access_token;
    } catch (err) {
        const detail = err.response?.data?.error_description || err.message;
        throw new Error(`Token refresh failed: ${detail}`);
    }
}

export function isOneDriveConnected() {
    return !!process.env.ONEDRIVE_REFRESH_TOKEN;
}
