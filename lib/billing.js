// Stripe billing: one-time Momentum credit packs + recurring subscription plans.
// Dormant until STRIPE_SECRET_KEY / STRIPE_WEBHOOK_SECRET are set.
const { supabase, supabaseAuth } = require('./store');

function getStripe() {
  const key = process.env.STRIPE_SECRET_KEY;
  if (!key) return null;
  try { return require('stripe')(key); } catch (e) { console.error('[billing] stripe sdk missing', e.message); return null; }
}

// Verify the caller's Supabase JWT and that they belong to (or admin over) the tenant.
async function authTenant(req, tenant_id) {
  const jwt = (req.headers.authorization || '').replace(/^Bearer\s+/i, '');
  const { data: { user } = {}, error } = await supabaseAuth.auth.getUser(jwt);
  if (error || !user) return { code: 401, error: 'unauthorized' };
  const { data: u } = await supabase.from('users').select('tenant_id, role').eq('id', user.id).single();
  if (!u || (u.tenant_id !== tenant_id && u.role !== 'we_admin')) return { code: 403, error: 'forbidden' };
  return { ok: true };
}

// POST /billing/checkout { tenant_id, pack_id }  — one-time Momentum credit pack.
async function stripeCheckout(req, res) {
  const stripe = getStripe();
  if (!stripe) return res.status(503).json({ error: 'stripe_not_configured' });
  const { tenant_id, pack_id, success_url, cancel_url } = req.body || {};
  if (!tenant_id || !pack_id) return res.status(400).json({ error: 'tenant_id and pack_id required' });
  const auth = await authTenant(req, tenant_id);
  if (!auth.ok) return res.status(auth.code).json({ error: auth.error });

  const { data: pack } = await supabase.from('credit_packs').select('*').eq('id', pack_id).eq('active', true).single();
  if (!pack) return res.status(400).json({ error: 'unknown_pack' });
  if (!pack.stripe_price_id) return res.status(400).json({ error: 'pack_missing_stripe_price' });
  try {
    const sess = await stripe.checkout.sessions.create({
      mode: 'payment',
      line_items: [{ price: pack.stripe_price_id, quantity: 1 }],
      metadata: { tenant_id, pack_id },
      success_url: success_url || 'https://www.intelligencemachine.ai/billing?credits=success',
      cancel_url: cancel_url || 'https://www.intelligencemachine.ai/billing?credits=cancel',
    });
    return res.json({ url: sess.url });
  } catch (e) {
    console.error('[billing] checkout failed', e.message);
    return res.status(500).json({ error: String(e.message || e) });
  }
}

// POST /billing/subscribe { tenant_id, tier }  — recurring subscription plan.
async function stripeSubscribe(req, res) {
  const stripe = getStripe();
  if (!stripe) return res.status(503).json({ error: 'stripe_not_configured' });
  const { tenant_id, tier, success_url, cancel_url } = req.body || {};
  if (!tenant_id || !tier) return res.status(400).json({ error: 'tenant_id and tier required' });
  const auth = await authTenant(req, tenant_id);
  if (!auth.ok) return res.status(auth.code).json({ error: auth.error });

  const { data: plan } = await supabase.from('plan_tiers').select('stripe_price_id, name').eq('tier', tier).single();
  if (!plan) return res.status(400).json({ error: 'unknown_tier' });
  if (!plan.stripe_price_id) return res.status(400).json({ error: 'tier_missing_stripe_price' });

  const { data: client } = await supabase.from('clients').select('stripe_customer_id, stripe_subscription_id, subscription_status').eq('tenant_id', tenant_id).maybeSingle();

  // Already subscribed → switch the plan on the existing subscription (proration), no new checkout.
  if (client && client.stripe_subscription_id && ['active', 'trialing', 'past_due'].includes(client.subscription_status)) {
    try {
      const sub = await stripe.subscriptions.retrieve(client.stripe_subscription_id);
      const updated = await stripe.subscriptions.update(sub.id, {
        items: [{ id: sub.items.data[0].id, price: plan.stripe_price_id }],
        proration_behavior: 'create_prorations',
        metadata: { tenant_id, tier },
      });
      await applySubscription(updated);
      return res.json({ switched: true });
    } catch (e) {
      console.error('[billing] plan switch failed', e.message);
      return res.status(500).json({ error: String(e.message || e) });
    }
  }

  try {
    const params = {
      mode: 'subscription',
      line_items: [{ price: plan.stripe_price_id, quantity: 1 }],
      metadata: { tenant_id, tier },
      subscription_data: { metadata: { tenant_id, tier } },
      success_url: success_url || 'https://www.intelligencemachine.ai/billing?sub=success',
      cancel_url: cancel_url || 'https://www.intelligencemachine.ai/billing?sub=cancel',
    };
    if (client && client.stripe_customer_id) params.customer = client.stripe_customer_id;
    const sess = await stripe.checkout.sessions.create(params);
    return res.json({ url: sess.url });
  } catch (e) {
    console.error('[billing] subscribe failed', e.message);
    return res.status(500).json({ error: String(e.message || e) });
  }
}

// Map a Stripe subscription object to a tier (via its price id) and apply it.
async function applySubscription(sub) {
  const priceId = sub.items && sub.items.data && sub.items.data[0] && sub.items.data[0].price && sub.items.data[0].price.id;
  let tier = sub.metadata && sub.metadata.tier;
  const tenant_id = sub.metadata && sub.metadata.tenant_id;
  if (priceId) {
    const { data: pt } = await supabase.from('plan_tiers').select('tier').eq('stripe_price_id', priceId).maybeSingle();
    if (pt) tier = pt.tier;
  }
  if (!tenant_id || !tier) return;
  await supabase.rpc('apply_subscription_tier', {
    p_tenant: tenant_id, p_tier: tier, p_subscription_id: sub.id, p_customer_id: sub.customer, p_status: sub.status,
  });
}

// POST /billing/webhook  (raw body; Stripe-signed).
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

  try {
    const o = event.data.object;
    switch (event.type) {
      case 'checkout.session.completed':
        if (o.mode === 'subscription') {
          // New subscription: fetch the full sub to get items/price, then apply.
          const sub = await stripe.subscriptions.retrieve(o.subscription);
          sub.metadata = { ...(sub.metadata || {}), ...(o.metadata || {}) };
          await applySubscription(sub);
        } else if (o.metadata && o.metadata.pack_id) {
          // One-time credit pack (idempotent by session id).
          await supabase.rpc('record_credit_purchase', { p_tenant: o.metadata.tenant_id, p_pack_id: o.metadata.pack_id, p_stripe_ref: o.id });
        }
        break;
      case 'customer.subscription.updated':
        await applySubscription(o);
        break;
      case 'customer.subscription.deleted':
        await supabase.rpc('set_subscription_status', { p_subscription_id: o.id, p_status: 'canceled' });
        break;
      case 'invoice.payment_failed':
        if (o.subscription) await supabase.rpc('set_subscription_status', { p_subscription_id: o.subscription, p_status: 'past_due' });
        break;
      default:
        break;
    }
  } catch (e) {
    console.error('[billing] webhook handler error', e.message);
    return res.status(500).end();
  }
  return res.json({ received: true });
}

module.exports = { stripeCheckout, stripeSubscribe, stripeWebhook };
