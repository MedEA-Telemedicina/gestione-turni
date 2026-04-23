// ============================================================
// Vercel Serverless Function: api/excel-proxy.js
// Proxy per esportazione su Excel Online via Microsoft Graph.
// Il refresh_token è salvato in Vercel KV (sostituisce Netlify Blobs).
// ============================================================

const { kv } = require('@vercel/kv');

const CLIENT_ID             = '879cce82-20a9-4a19-8090-da52783d4356';
const TENANT_ID             = '9517c96a-8492-4ea0-a5e4-c791b8006c61';
const EXCEL_SHARING_ENCODED = 'u!aHR0cHM6Ly9tZWRlYWRpZ2l0YWxoZWFsdGgtbXkuc2hhcmVwb2ludC5jb20vOng6L2cvcGVyc29uYWwvYV9tYXJpbm9fbWVkLWVhX2l0L0lRQ0lUeTBteFZPSVRycmpocC1kS2p0dUFjaDRaUDVObndsdnlOTkJzaWQ5VktzP2U9cjRYWWFC';
const EXCEL_WORKSHEET       = 'TURNI DEFINITIVI';
const EXCEL_START_ROW       = 18;

// ---- helper ----
function colToLetter(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(65 + (n - 1) % 26) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function setCors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
}

// ---- main handler ----
module.exports = async function handler(req, res) {

  setCors(res);

  // Pre-flight CORS
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Metodo non consentito.' });
  }

  // Leggi refresh token da Vercel KV
  const refreshToken = await kv.get('refresh_token');

  if (!refreshToken) {
    return res.status(500).json({
      error: 'Refresh token non trovato. Esegui prima get-token.html per autenticarti.',
    });
  }

  try {
    // ------ 1. Parse body ------
    const { data } = req.body || {};
    if (!Array.isArray(data) || !data.length) {
      throw new Error('Nessun dato ricevuto dal client.');
    }

    // ------ 2. Ottieni access token dal refresh token ------
    const tokenResp = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id:     CLIENT_ID,
          client_secret: process.env.GRAPH_CLIENT_SECRET || '',
          grant_type:    'refresh_token',
          refresh_token: refreshToken,
          scope:         'https://graph.microsoft.com/Files.ReadWrite offline_access',
        }),
      }
    );
    const tokenData = await tokenResp.json();

    if (!tokenData.access_token) {
      const desc = tokenData.error_description || tokenData.error || JSON.stringify(tokenData);
      throw new Error('Autenticazione Microsoft fallita: ' + desc);
    }

    const token = tokenData.access_token;

    // Aggiorna il refresh token in KV se Microsoft ne ha emesso uno nuovo (rotazione)
    if (tokenData.refresh_token) {
      await kv.set('refresh_token', tokenData.refresh_token);
      console.log('[excel-proxy] Refresh token aggiornato in Vercel KV.');
    }

    // ------ 3. Risolvi il link condiviso → driveId + itemId ------
    const resolveResp = await fetch(
      `https://graph.microsoft.com/v1.0/shares/${EXCEL_SHARING_ENCODED}/driveItem?$select=id,parentReference`,
      { headers: { Authorization: 'Bearer ' + token } }
    );

    if (!resolveResp.ok) {
      const err = await resolveResp.json().catch(() => ({}));
      throw new Error('Impossibile risolvere il file condiviso: ' + (err?.error?.message || resolveResp.status));
    }

    const resolveData = await resolveResp.json();
    const itemId  = resolveData.id;
    const driveId = resolveData.parentReference?.driveId;
    if (!itemId || !driveId) throw new Error('Risposta inattesa dalla risoluzione del file condiviso.');

    // ------ 4. Calcola range ------
    const numRows  = data.length;
    const numCols  = data[0].length;
    const endCol   = colToLetter(numCols);
    const endRow   = EXCEL_START_ROW + numRows - 1;
    const rangeAddr = `A${EXCEL_START_ROW}:${endCol}${endRow}`;

    const baseUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(EXCEL_WORKSHEET)}')`;

    // ------ 5. Cancella range precedente ------
    await fetch(`${baseUrl}/range(address='A${EXCEL_START_ROW}:Z${endRow + 10}')/clear`, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ applyTo: 'Contents' }),
    });

    // ------ 6. Scrivi nuovi dati ------
    const writeResp = await fetch(`${baseUrl}/range(address='${rangeAddr}')`, {
      method: 'PATCH',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: data }),
    });

    if (!writeResp.ok) {
      const errData = await writeResp.json().catch(() => ({}));
      throw new Error(errData?.error?.message || 'HTTP ' + writeResp.status);
    }

    // ------ 7. Risposta OK ------
    return res.status(200).json({ success: true, range: rangeAddr });

  } catch (err) {
    console.error('[excel-proxy] Errore:', err);
    return res.status(500).json({ error: err.message || String(err) });
  }
};
