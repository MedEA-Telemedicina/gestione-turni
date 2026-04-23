// ============================================================
// Vercel Serverless Function: api/token-exchange.js
// Scambia il codice OAuth lato server, salva il refresh token
// in Vercel KV (sostituisce Netlify Blobs).
// ============================================================

const { kv } = require('@vercel/kv');

const CLIENT_ID = '879cce82-20a9-4a19-8090-da52783d4356';
const TENANT_ID = '9517c96a-8492-4ea0-a5e4-c791b8006c61';

function setCors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
}

module.exports = async function handler(req, res) {

  setCors(res);

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Metodo non consentito.' });
  }

  const { code, codeVerifier, redirectUri } = req.body || {};

  if (!code || !codeVerifier || !redirectUri) {
    return res.status(400).json({
      error: 'Parametri mancanti: code, codeVerifier, redirectUri.',
    });
  }

  try {
    const tokenResp = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id:     CLIENT_ID,
          client_secret: process.env.GRAPH_CLIENT_SECRET || '',
          grant_type:    'authorization_code',
          code,
          redirect_uri:  redirectUri,
          code_verifier: codeVerifier,
          scope:         'https://graph.microsoft.com/Files.ReadWrite offline_access',
        }),
      }
    );

    const data = await tokenResp.json();

    if (tokenResp.ok && data.refresh_token) {
      // Salva il refresh token in Vercel KV
      await kv.set('refresh_token', data.refresh_token);
      console.log('[token-exchange] Refresh token salvato in Vercel KV.');
    }

    // Rispondiamo senza esporre il refresh token al browser
    return res.status(tokenResp.ok ? 200 : 400).json(
      tokenResp.ok
        ? { success: true, message: 'Token salvato correttamente sul server.' }
        : data
    );

  } catch (err) {
    console.error('[token-exchange] Errore:', err);
    return res.status(500).json({ error: 'Errore server: ' + err.message });
  }
};
