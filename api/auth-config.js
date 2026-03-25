import * as msal from '@azure/msal-node';

// ─── MSAL Config ──────────────────────────────────────────────────────────────

const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID || 'common'}`,
    },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

// In-memory token cache (persists for process lifetime)
// For production: swap with Redis or a database store
let cachedTokens = {
    accessToken: process.env.ONEDRIVE_ACCESS_TOKEN || null,
    refreshToken: process.env.ONEDRIVE_REFRESH_TOKEN || null,
    expiresOn: null,
};

// ─── Get Auth Code URL (step 1 of OAuth flow) ─────────────────────────────────

export async function getAuthCodeUrl(redirectUri) {
    const authCodeUrlParams = {
        scopes: [
            'https://graph.microsoft.com/Files.ReadWrite',
            'https://graph.microsoft.com/Files.ReadWrite.All',
            'offline_access',
        ],
        redirectUri,
        prompt: 'consent',
    };
    return await cca.getAuthCodeUrl(authCodeUrlParams);
}

// ─── Exchange Code for Tokens (step 2 of OAuth flow) ─────────────────────────

export async function exchangeCodeForTokens(code, redirectUri) {
    const tokenRequest = {
        code,
        scopes: [
            'https://graph.microsoft.com/Files.ReadWrite',
            'https://graph.microsoft.com/Files.ReadWrite.All',
            'offline_access',
        ],
        redirectUri,
    };

    const response = await cca.acquireTokenByCode(tokenRequest);

    // Cache in memory
    cachedTokens = {
        accessToken: response.accessToken,
        refreshToken: response.account?.refreshToken || null,
        expiresOn: response.expiresOn,
    };

    console.log('✅ Tokens acquired and cached');
    return cachedTokens;
}

// ─── Get Valid Access Token (auto-refresh if expired) ─────────────────────────

export async function getAccessToken() {
    const now = new Date();
    const expiresOn = cachedTokens.expiresOn ? new Date(cachedTokens.expiresOn) : null;
    const isExpired = !expiresOn || expiresOn <= now;

    // Token is still valid
    if (cachedTokens.accessToken && !isExpired) {
        return cachedTokens.accessToken;
    }

    // Try refresh token
    if (cachedTokens.refreshToken) {
        try {
            console.log('🔄 Refreshing access token...');
            // MSAL handles refresh internally via acquireTokenSilent
            // We use the cached account or fall back to refresh token grant
            const silentRequest = {
                scopes: [
                    'https://graph.microsoft.com/Files.ReadWrite',
                    'https://graph.microsoft.com/Files.ReadWrite.All',
                    'offline_access',
                ],
                forceRefresh: true,
            };

            // Try to get accounts
            const accounts = await cca.getTokenCache().getAllAccounts();
            if (accounts.length > 0) {
                silentRequest.account = accounts[0];
                const response = await cca.acquireTokenSilent(silentRequest);
                cachedTokens.accessToken = response.accessToken;
                cachedTokens.expiresOn = response.expiresOn;
                console.log('✅ Token refreshed');
                return cachedTokens.accessToken;
            }
        } catch (err) {
            console.warn('⚠️  Silent refresh failed:', err.message);
        }
    }

    // Fall back to env token (set manually via Vercel env vars after OAuth)
    if (process.env.ONEDRIVE_ACCESS_TOKEN) {
        console.log('ℹ️  Using ONEDRIVE_ACCESS_TOKEN from environment');
        return process.env.ONEDRIVE_ACCESS_TOKEN;
    }

    throw new Error(
        'No valid token available. Complete OAuth flow at /api/oauth/authorize-url'
    );
}