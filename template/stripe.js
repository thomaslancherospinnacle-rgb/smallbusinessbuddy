/* ═══════════════════════════════════════════════════════════════
   {{COMPANY_SHORT}} Stripe + Email Worker — Cloudflare Worker
   ═══════════════════════════════════════════════════════════════

   ENV VARS NEEDED (set in Cloudflare Worker dashboard):
     STRIPE_SECRET_KEY   = sk_live_xxxxx  (or sk_test_xxxxx)
     RESEND_API_KEY      = re_xxxxx
     SITE_URL            = https://{{COMPANY_SLUG}}.fit  (no trailing slash)

   ENDPOINTS:
     POST /email          → send email via Resend (existing)
     POST /checkout       → create $100 deposit Stripe Checkout session
     POST /invoice        → create custom-amount Stripe Payment Link
     POST /webhook        → Stripe webhook handler (optional, for paid status)
     GET  /               → health check

   STRIPE SETUP:
     1. stripe.com → Dashboard → API Keys → copy Secret Key
     2. Add as env var STRIPE_SECRET_KEY
     3. In Stripe Dashboard → Products → create "Inspection Deposit" product ($100)
        (optional — the worker creates ad-hoc prices so no product needed)
     4. For webhooks (optional): Stripe → Webhooks → Add endpoint → 
        https://your-worker.workers.dev/webhook
        Events: checkout.session.completed, payment_link.payment_link.created

   ═══════════════════════════════════════════════════════════════ */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    // ── CORS preflight ──
    if (request.method === 'OPTIONS') {
      return cors(null, 204);
    }

    // ── Health check ──
    if (request.method === 'GET' && path === '/') {
      return cors(json({ status: 'ok', worker: '{{COMPANY_SHORT}} Stripe+Email', version: '2.0' }));
    }

    if (request.method !== 'POST') {
      return cors(json({ error: 'POST only' }, 405));
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return cors(json({ error: 'Invalid JSON' }, 400));
    }

    // ── Route ──
    if (path === '/email')    return handleEmail(body, env);
    if (path === '/checkout') return handleCheckout(body, env);
    if (path === '/invoice')  return handleInvoice(body, env);
    if (path === '/webhook')  return handleWebhook(request, env);

    return cors(json({ error: 'Not found' }, 404));
  }
};

/* ══════════════════════════════════════════
   EMAIL  (Resend — same as before)
   ══════════════════════════════════════════ */
async function handleEmail(data, env) {
  const { from, to, subject, body, replyTo } = data;
  if (!to)    return cors(json({ error: 'Missing "to"' }, 400));
  if (!from)  return cors(json({ error: 'Missing "from"' }, 400));
  if (!subject && !body) return cors(json({ error: 'Missing subject/body' }, 400));

  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${env.RESEND_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from,
      to: [to],
      subject: subject || '(no subject)',
      text: body || '',
      reply_to: replyTo || from,
    }),
  });

  const result = await res.json();
  if (res.ok) return cors(json({ success: true, id: result.id }));
  return cors(json({ error: result.message || 'Resend error', details: result }, 400));
}

/* ══════════════════════════════════════════
   CHECKOUT  — $100 Deposit
   Body: { customerName, customerEmail, service, appointmentDate, appointmentTime, appointmentId }
   Returns: { url } — redirect client to this Stripe Checkout URL
   ══════════════════════════════════════════ */
async function handleCheckout(data, env) {
  const {
    customerName,
    customerEmail,
    service = 'Plumbing Inspection',
    appointmentDate = '',
    appointmentTime = '',
    appointmentId = '',
  } = data;

  if (!customerEmail) return cors(json({ error: 'Missing customerEmail' }, 400));

  const siteUrl = env.SITE_URL || 'https://{{COMPANY_SLUG}}.fit';

  // Build metadata so you can identify this booking in the webhook
  const metadata = {
    appointmentId: String(appointmentId),
    service,
    appointmentDate,
    appointmentTime,
    customerName: customerName || '',
  };

  // Create a Stripe Checkout Session
  const params = new URLSearchParams({
    'payment_method_types[]': 'card',
    'mode': 'payment',
    'customer_email': customerEmail,
    'success_url': `${siteUrl}/booking-success.html?session_id={CHECKOUT_SESSION_ID}&appt=${encodeURIComponent(appointmentId)}`,
    'cancel_url': `${siteUrl}/#book`,
    'line_items[0][price_data][currency]': 'usd',
    'line_items[0][price_data][unit_amount]': '{{DEPOSIT_CENTS}}',  // $100.00 in cents
    'line_items[0][price_data][product_data][name]': `Inspection Deposit — ${service}`,
    'line_items[0][price_data][product_data][description]': appointmentDate
      ? `${appointmentDate}${appointmentTime ? ' at ' + appointmentTime : ''} · Non-refundable travel & inspection fee`
      : 'Non-refundable travel & inspection fee',
    'line_items[0][quantity]': '1',
    'payment_intent_data[description]': `{{COMPANY_SHORT}} deposit: ${service} on ${appointmentDate}`,
  });

  // Append metadata
  for (const [k, v] of Object.entries(metadata)) {
    params.append(`metadata[${k}]`, v);
    params.append(`payment_intent_data[metadata][${k}]`, v);
  }

  const res = await stripePost('checkout/sessions', params, env);
  const result = await res.json();

  if (res.ok) {
    return cors(json({ success: true, url: result.url, sessionId: result.id }));
  }
  return cors(json({ error: result.error?.message || 'Stripe error', details: result }, 400));
}

/* ══════════════════════════════════════════
   INVOICE  — Custom amount Payment Link
   Body: { amount, description, customerName, customerEmail, appointmentId, jobId }
   Returns: { url, qr } — shareable payment link + QR PNG url
   ══════════════════════════════════════════ */
async function handleInvoice(data, env) {
  const {
    amount,           // dollars, e.g. 385.00
    description = '{{INDUSTRY_TITLE}} Services',
    customerName = '',
    customerEmail = '',
    appointmentId = '',
    jobId = '',
  } = data;

  if (!amount || isNaN(parseFloat(amount))) {
    return cors(json({ error: 'Missing or invalid amount' }, 400));
  }

  const cents = Math.round(parseFloat(amount) * 100);
  if (cents < 50) return cors(json({ error: 'Minimum invoice is $0.50' }, 400));

  // Step 1: Build a descriptive product name (Stripe Prices API only takes [name], no [description])
  const jobRef = jobId || appointmentId;
  const productName = [
    description,
    customerName ? `· ${customerName}` : '',
    jobRef      ? `· Job #${jobRef}` : '',
  ].filter(Boolean).join(' ');

  const priceParams = new URLSearchParams({
    'currency': 'usd',
    'unit_amount': String(cents),
    'product_data[name]': productName.slice(0, 250), // Stripe max 250 chars
  });

  const priceRes = await stripePost('prices', priceParams, env);
  const price = await priceRes.json();
  if (!priceRes.ok) {
    return cors(json({ error: price.error?.message || 'Stripe price creation failed', stripe: price }, 400));
  }

  // Step 2: Create a Payment Link from that Price
  // NOTE: invoice_creation is omitted — requires Stripe invoicing add-on
  const linkParams = new URLSearchParams({
    'line_items[0][price]': price.id,
    'line_items[0][quantity]': '1',
    'metadata[appointmentId]': String(appointmentId),
    'metadata[jobId]': String(jobId),
    'metadata[customerName]': customerName,
    'metadata[customerEmail]': customerEmail,
    'metadata[type]': 'job_invoice',
    'after_completion[type]': 'redirect',
    'after_completion[redirect][url]': `${env.SITE_URL || 'https://{{COMPANY_SLUG}}.fit'}/booking-complete/index.html`,
  });

  const linkRes = await stripePost('payment_links', linkParams, env);
  const link = await linkRes.json();
  if (!linkRes.ok) {
    return cors(json({ error: link.error?.message || 'Stripe payment link failed', stripe: link }, 400));
  }

  // Step 3: QR code via free QR API (no key needed)
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(link.url)}`;

  return cors(json({
    success: true,
    url: link.url,
    qrUrl,
    priceId: price.id,
    linkId: link.id,
    amount: cents / 100,
    description,
  }));
}

/* ══════════════════════════════════════════
   WEBHOOK  — Stripe → update GAS (optional)
   ══════════════════════════════════════════ */
async function handleWebhook(request, env) {
  // If you add STRIPE_WEBHOOK_SECRET env var, this will verify the signature.
  // Without it, it still processes but isn't signature-verified.
  const payload = await request.text();

  let event;
  try {
    event = JSON.parse(payload);
  } catch {
    return new Response('Bad payload', { status: 400 });
  }

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;
    const meta = session.metadata || {};
    console.log(`Deposit paid: appointmentId=${meta.appointmentId}, customer=${session.customer_email}, amount=${session.amount_total}`);
    // If you want to auto-update your GAS sheet:
    // await notifyGAS(env, { action: 'markDepositPaid', appointmentId: meta.appointmentId, amount: session.amount_total / 100 });
  }

  if (event.type === 'payment_intent.succeeded') {
    const pi = event.data.object;
    const meta = pi.metadata || {};
    if (meta.type === 'job_invoice') {
      console.log(`Invoice paid: jobId=${meta.jobId}, amount=${pi.amount / 100}`);
      // await notifyGAS(env, { action: 'markInvoicePaid', jobId: meta.jobId, amount: pi.amount / 100 });
    }
  }

  return new Response(JSON.stringify({ received: true }), { status: 200, headers: { 'Content-Type': 'application/json' } });
}

/* ══════════════════════════════════════════
   HELPERS
   ══════════════════════════════════════════ */
async function stripePost(endpoint, params, env) {
  return fetch(`https://api.stripe.com/v1/${endpoint}`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${env.STRIPE_SECRET_KEY}`,
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params.toString(),
  });
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function cors(response, status) {
  if (response === null) {
    return new Response(null, {
      status: status || 204,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      },
    });
  }
  // Must clone before reading headers/body — consuming either on the original
  // exhausts the stream and causes the response to arrive empty at the client.
  const clone = response.clone();
  const headers = new Headers(clone.headers);
  headers.set('Access-Control-Allow-Origin', '*');
  headers.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  headers.set('Access-Control-Allow-Headers', 'Content-Type');
  return new Response(clone.body, { status: clone.status, headers });
}
