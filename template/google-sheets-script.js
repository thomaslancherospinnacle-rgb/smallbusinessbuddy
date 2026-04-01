/* ═══════════════════════════════════════════════════════════════
   {{COMPANY_SHORT}} — GAS Proxy Worker
   {{GAS_WORKER_URL}}

   ENV VARS (set in Cloudflare Worker dashboard → Settings → Variables):
     GAS_URL  =  https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec

   ENDPOINTS:
     GET  /?action=&...  → forwarded to GAS as-is
     GET  /?action=getDashboard         → forwarded to GAS as-is
     GET  /?action=getRevenue           → forwarded to GAS as-is
     GET  /?action=ping                 → forwarded to GAS as-is
     POST /  { action: 'appointment', ... }  → forwarded to GAS
     GET  /health  → local health check (does NOT hit GAS)

   HOW TO UPDATE GAS URL:
     Cloudflare Dashboard → Workers & Pages → google-sheets-script
     → Settings → Variables → edit GAS_URL
     No code change, no GitHub commit needed.

   ═══════════════════════════════════════════════════════════════ */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // ── CORS preflight ──
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }

    // ── Local health check — does NOT proxy to GAS ──
    if (url.pathname === '/health') {
      return respond({
        status: 'ok',
        worker: '{{COMPANY_SHORT}} GAS Proxy',
        gasConfigured: !!env.GAS_URL,
        timestamp: new Date().toISOString(),
      });
    }

    // ── Root + /api both proxy to GAS ──
    if (url.pathname === '/' || url.pathname === '/api') {
      const gasUrl = env.GAS_URL;

      if (!gasUrl) {
        return respond({
          error: 'GAS_URL environment variable not set. Add it in the Cloudflare Worker dashboard → Settings → Variables.'
        }, 500);
      }

      try {
        let gasResponse;

        // ── GET: forward all query params to GAS ──
        if (request.method === 'GET') {
          const target = new URL(gasUrl);
          // Copy every query param from the incoming request to GAS URL
          url.searchParams.forEach((val, key) => {
            target.searchParams.set(key, val);
          });
          gasResponse = await fetch(target.toString(), {
            method: 'GET',
            headers: { 'Accept': 'application/json' },
            // GAS redirects — follow them
            redirect: 'follow',
          });
        }

        // ── POST: forward JSON body to GAS ──
        // GAS doPost reads e.postData.contents regardless of Content-Type.
        // We send as text/plain to avoid GAS CORS preflight issues on redirect.
        else if (request.method === 'POST') {
          const body = await request.text();
          gasResponse = await fetch(gasUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'text/plain;charset=utf-8',
            },
            body,
            redirect: 'follow',
          });
        }

        else {
          return respond({ error: 'Only GET and POST are supported' }, 405);
        }

        // Forward the GAS response back to the client with CORS headers
        const responseText = await gasResponse.text();
        return new Response(responseText, {
          status: gasResponse.ok ? 200 : gasResponse.status,
          headers: {
            'Content-Type': 'application/json',
            ...corsHeaders(),
          },
        });

      } catch (err) {
        return respond({
          error: 'Worker failed to reach GAS',
          detail: err.message,
        }, 502);
      }
    }

    // ── 404 for anything else ──
    return respond({ error: 'Not found', path: url.pathname }, 404);
  }
};

/* ── Helpers ── */

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}

function respond(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'Content-Type': 'application/json',
      ...corsHeaders(),
    },
  });
}
