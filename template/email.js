/* ═══════════════════════════════════════════════════════════════
   {{COMPANY_SHORT}} Email Worker — Cloudflare Worker + Resend
   ═══════════════════════════════════════════════════════════════
   
   SETUP:
   1. Go to dash.cloudflare.com → Workers & Pages → Create
   2. Paste this code
   3. Add environment variable:  RESEND_API_KEY = re_xxxxxxxx
   4. Deploy
   5. Copy the worker URL into admin.html Settings → Email Worker URL
   
   RESEND SETUP:
   1. resend.com → Add your domain ({{COMPANY_SLUG}}.fit)
   2. Add the DNS records Resend gives you (DKIM, SPF, MX) in Cloudflare
   3. Verify domain
   4. Create API key → paste as RESEND_API_KEY env var
   
   ImprovMX SETUP (for RECEIVING):
   1. improvmx.com → Add domain ({{COMPANY_SLUG}}.fit)
   2. Add MX records in Cloudflare DNS
   3. Set forwarding: *@{{COMPANY_SLUG}}.fit → yourgmail@gmail.com
   
   That's it. Outbound goes through Resend, inbound through ImprovMX → Gmail.
   
   ═══════════════════════════════════════════════════════════════ */

export default {
  async fetch(request, env) {
    // CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type',
        },
      });
    }

    if (request.method !== 'POST') {
      return jsonRes({ error: 'POST only' }, 405);
    }

    try {
      const data = await request.json();
      const { from, to, subject, body, replyTo } = data;

      // Validate
      if (!to) return jsonRes({ error: 'Missing "to"' }, 400);
      if (!from) return jsonRes({ error: 'Missing "from"' }, 400);
      if (!subject && !body) return jsonRes({ error: 'Missing subject/body' }, 400);

      // Send via Resend
      const resendRes = await fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${env.RESEND_API_KEY}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          from: from,          // e.g. "Support <support@{{COMPANY_SLUG}}.fit>"
          to: [to],
          subject: subject || '(no subject)',
          text: body || '',
          reply_to: replyTo || from,
        }),
      });

      const result = await resendRes.json();

      if (resendRes.ok) {
        return jsonRes({ success: true, id: result.id, to, subject });
      } else {
        return jsonRes({ error: result.message || 'Resend error', details: result }, 400);
      }

    } catch (err) {
      return jsonRes({ error: err.message }, 500);
    }
  },
};

function jsonRes(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*',
    },
  });
}
