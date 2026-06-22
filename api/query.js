// Vercel Serverless Function: /api/query.js
// This runs on Vercel's servers (not the browser), so the secret stays hidden
// and there's no CORS problem with Azure AD.

export default async function handler(req, res) {
  // Allow the chatbot (from GitHub Pages) to call this proxy
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // Handle CORS preflight
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // Config from environment variables (set in Vercel dashboard - NOT in code)
  const TENANT_ID = process.env.PBI_TENANT_ID;
  const CLIENT_ID = process.env.PBI_CLIENT_ID;
  const CLIENT_SECRET = process.env.PBI_CLIENT_SECRET;
  const WORKSPACE_ID = process.env.PBI_WORKSPACE_ID;
  const DATASET_ID = process.env.PBI_DATASET_ID;

  try {
    const { dax } = req.body;
    if (!dax) {
      return res.status(400).json({ error: 'Missing dax query' });
    }

    // Step 1: Get a token from Azure AD (server-side, no CORS issue here)
    const tokenResp = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'client_credentials',
          client_id: CLIENT_ID,
          client_secret: CLIENT_SECRET,
          scope: 'https://analysis.windows.net/powerbi/api/.default'
        })
      }
    );

    if (!tokenResp.ok) {
      const err = await tokenResp.text();
      return res.status(500).json({ error: 'Token failed', detail: err });
    }

    const tokenData = await tokenResp.json();
    const accessToken = tokenData.access_token;

    // Step 2: Run the DAX query against Power BI
    const queryResp = await fetch(
      `https://api.powerbi.com/v1.0/myorg/groups/${WORKSPACE_ID}/datasets/${DATASET_ID}/executeQueries`,
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          queries: [{ query: dax }],
          serializerSettings: { includeNulls: true }
        })
      }
    );

    if (!queryResp.ok) {
      const err = await queryResp.text();
      return res.status(500).json({ error: 'Query failed', detail: err });
    }

    const queryData = await queryResp.json();
    const rows = queryData.results[0].tables[0].rows;

    return res.status(200).json({ rows });

  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
