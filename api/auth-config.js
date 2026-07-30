// Minimal OneDrive OAuth — used ONCE to obtain the initial refresh token.
// After the token is seeded into GitHub Secrets, the GitHub Action handles
// all subsequent token rotation. The CMS itself never calls getAccessToken().

import axios from 'axios';

const TENANT_ID = process.env.ONEDRIVE_TENANT_ID || process.env.AZURE_TENANT_ID || 'common';
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const SCOPES = 'https://graph.microsoft.com/Files.ReadWrite offline_access';

export function getAuthCodeUrl(redirectUri) {
    const state = Buffer.from(redirectUri).toString('base64');
    const params = new URLSearchParams({
        client_id: process.env.AZURE_CLIENT_ID,
        response_type: 'code',
        redirect_uri: redirectUri,
        scope: SCOPES,
        response_mode: 'query',
        prompt: 'consent',
        state,
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
    return res.data; // { access_token, refresh_token, expires_in, ... }
}
