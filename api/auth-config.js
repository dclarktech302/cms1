import axios from 'axios';

// ─── Constants ────────────────────────────────────────────────────────────────

const SCOPES = [
    'https://graph.microsoft.com/Files.ReadWrite',
    'https://graph.microsoft.com/Files.ReadWrite.All',
    'offline_access',
].join(' ');

function tenantUrl() {
    const tenant = process.env.AZURE_TENANT_ID || 'common';
    return `https://login.microsoftonline.com/${tenant}`;
}

// ─── In-memory token cache ────────────────────────────────────────────────────

let cachedTokens = {
    accessToken: process.env.ONEDRIVE_ACCESS_TOKEN || null,
    refreshToken: process.env.ONEDRIVE_REFRESH_TOKEN || null,
    expiresAt: null,
};

// ─── Step 1 — Build authorization URL ────────────────────────────────────────

export function getAuthCodeUrl(redirectUri) {
    const params = new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        response_type: 'code',
        redirect_uri: redirectUri,
        scope: SCOPES,
        response_mode: 'query',
        prompt: 'consent',
    });
    return `${tenantUrl()}/oauth2/v2.0/authorize?${params.toString()}`;
}

// ─── Step 2 — Exchange auth code for tokens (raw HTTP) ───────────────────────
// MSAL never exposes the refresh token on the response object — it hides it
// in its internal cache. We call the token endpoint directly instead.

export async function exchangeCodeForTokens(code, redirectUri) {
    const body = new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        code,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
        scope: SCOPES,
    });

    const response = await axios.post(
        `${tenantUrl()}/oauth2/v2.0/token`,
        body.toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token, expires_in } = response.data;

    cachedTokens = {
        accessToken: access_token,
        refreshToken: refresh_token || null,
        expiresAt: new Date(Date.now() + expires_in * 1000),
    };

    console.log('Tokens acquired:', {
        hasAccessToken: !!access_token,
        hasRefreshToken: !!refresh_token,
        expiresAt: cachedTokens.expiresAt,
    });

    return {
        accessToken: access_token,
        refreshToken: refresh_token || null,
        expiresOn: cachedTokens.expiresAt,
    };
}

// ─── Refresh using refresh token ──────────────────────────────────────────────

async function refreshAccessToken() {
    if (!cachedTokens.refreshToken) {
        throw new Error('No refresh token available — re-run OAuth flow');
    }

    console.log('Refreshing access token...');

    const body = new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        refresh_token: cachedTokens.refreshToken,
        grant_type: 'refresh_token',
        scope: SCOPES,
    });

    const response = await axios.post(
        `${tenantUrl()}/oauth2/v2.0/token`,
        body.toString(),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token, expires_in } = response.data;

    cachedTokens = {
        accessToken: access_token,
        refreshToken: refresh_token || cachedTokens.refreshToken,
        expiresAt: new Date(Date.now() + expires_in * 1000),
    };

    console.log('Token refreshed, expires:', cachedTokens.expiresAt);
    return cachedTokens.accessToken;
}

// ─── Get a valid access token (auto-refresh if needed) ────────────────────────

export async function getAccessToken() {
    const now = new Date();
    const bufferMs = 60 * 1000;
    const isExpired =
        !cachedTokens.expiresAt ||
        cachedTokens.expiresAt.getTime() - bufferMs <= now.getTime();

    if (cachedTokens.accessToken && !isExpired) {
        return cachedTokens.accessToken;
    }

    if (cachedTokens.refreshToken) {
        return await refreshAccessToken();
    }

    if (process.env.ONEDRIVE_ACCESS_TOKEN) {
        console.log('Using ONEDRIVE_ACCESS_TOKEN from environment');
        return process.env.ONEDRIVE_ACCESS_TOKEN;
    }

    throw new Error(
        'No valid token available. Visit /api/oauth/authorize-url to re-authenticate.'
    );
}