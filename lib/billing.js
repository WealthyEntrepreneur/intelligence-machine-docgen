// Stripe checkout + webhook for Momentum Profile credit packs.
// Stays dormant until STRIPE_SECRET_KEY / STRIPE_WEBHOOK_SECRET are set.
const { supabase } = require('./store');

function getStripe() {
  const key = process.env.STRIPE_SECRET_KEY;
  if (!key) return null;
  try { return require('stripe')(key); } catch (e) { console.error('[billing] stripe sdk missing', e.message); return null; }
}

// POST /billing/checkout { tenant_id, pack_id, success_url?, cancel_url? }
// Auth: the caller's Supabase JWT (Bearer) — verified here, so the browser never holds a service secret.
async function stripeCheckout(req, res) {
  const stripe = getStripe();
  if (!stripe) return res.status(503).json({ error: 'stripe_not_configured' });

  const jwt = (req.headers.authorization || '').replace(/^Bearer\s+/i, '');
  const { data: { user } = {}, error: authErr } = await supabase.auth.getUser(jwt);
  if (authErr || !user) return res.status(401).json({ error: 'unauthorized' });

  const { tenant_id, pack_id, success_url, cancel_url } = req.body || {};
  if (!tenant_id || !pack_id) return res.status(400).json({ error: 'tenant_id and pack_id required' });

  // Confirm the caller belongs to this tenant.
  const { data: u } = await supabase.from('users').select('tenant_id, role').eq('id', user.id).single();
  if (!u || (u.tenant_id !== tenant_id && u.role !== 'we_admin')) return res.status(403).json({ error: 'forbidden' });

  const { data: pack } = await supabase.from('credit_packs').select('*').eq('id', pack_id).eq('active', true).single();
  if (!pack) return res.status(400).json({ error: 'unknown_pack' });
  if (!pack.stripe_price_id) return res.status(400).json({ error: 'pack_missing_stripe_price' });

  try {
    const sess = await stripe.checkout.sessions.create({
      mode: 'payment',
      line_items: [{ price: pack.stripe_price_id, quantity: 1 }],
      metadata: { tenant_id, pack_id },
      success_url: success_url || 'https://www.intelligencemachine.ai/settings?credits=success',
      cancel_url: cancel_url || 'https://www.intelligencemachine.ai/settings?credits=cancel',
    });
    return res.json({ url: sess.url });
  } catch (e) {
    console.error('[billing] checkout failed', e.message);
    return res.status(500).json({ error: String(e.message || e) });
  }
}

// POST /billing/webhook  (raw body; Stripe-signed). Idempotent credit on checkout.session.completed.
async function stripeWebhook(req, res) {
  const stripe = getStripe();
  const whSecret = process.env.STRIPE_WEBHOOK_SECRET;
  if (!stripe || !whSecret) return res.status(503).end();

  let event;
  try {
    event = stripe.webhooks.constructEvent(req.body, req.headers['stripe-signature'], whSecret);
  } catch (e) {
    return res.status(400).send(`Webhook signature error: ${e.message}`);
  }

  if (event.type === 'checkout.session.completed') {
    const s = event.data.object;
    const tenant_id = s.metadata && s.metadata.tenant_id;
    const pack_id = s.metadata && s.metadata.pack_id;
    if (tenant_id && pack_id) {
      try {
        await supabase.rpc('record_credit_purchase', { p_tenant: tenant_id, p_pack_id: pack_id, p_stripe_ref: s.id });
      } catch (e) {
        console.error('[billing] record_credit_purchase failed', e.message);
        return res.status(500).end();
      }
    }
  }
  return res.json({ received: true });
}

module.exports = { stripeCheckout, stripeWebhook };
