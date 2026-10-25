/* XLSX Comparator — free trial + Microsoft Store subscription
 *
 * Model: first FREE_COMPARISONS comparisons are free; after that a monthly
 * subscription (a Microsoft Store "Subscription" add-on) is required.
 * Payment is handled entirely by the Microsoft Store (card / other Store payment
 * methods). The app never sees or stores payment details.
 *
 * >>> Edit CONFIG below to match what you create in Partner Center. <<<
 */
const CONFIG = {
  ENABLED: true,                       // set false to disable all limits
  FREE_COMPARISONS: 5,                 // free comparison runs before the paywall
  PRODUCT_IDS: ['xlsx_pro_monthly'],   // MUST equal the Product ID of your subscription add-on
  BILLING_URL: 'https://store.microsoft.com/billing', // Microsoft Store billing payment method for the Digital Goods API
  STORE_PAGE: 'https://apps.microsoft.com/detail/9N4DHJQ04RM9',
  OFFLINE_GRACE_DAYS: 3,               // keep Premium unlocked this long if the Store can't be reached
  APP_NAME: 'XLSX Comparator'
};

const Monetization = (function () {
  const USAGE_KEY = 'xc_usage_v1';
  const ENT_KEY = 'xc_entitlement_v1';
  const SALT = 'xc-local-integrity';
  let premium = false;
  let busy = false;
  let priceText = '';

  /* ---------- tiny integrity check (stops casual edits, NOT a security boundary) ---------- */
  function sign(str) {
    let h = 5381;
    const s = str + SALT;
    for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
    return String(h >>> 0);
  }
  function readStore(key) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const o = JSON.parse(raw);
      const sig = o.sig; delete o.sig;
      return sign(JSON.stringify(o)) === sig ? o : null;
    } catch (e) { return null; }
  }
  function writeStore(key, obj) {
    try {
      const o = Object.assign({}, obj);
      o.sig = sign(JSON.stringify(obj));
      localStorage.setItem(key, JSON.stringify(o));
    } catch (e) { /* storage unavailable */ }
  }

  /* ---------- usage counter ---------- */
  function getUsed() {
    const u = readStore(USAGE_KEY);
    // A tampered/corrupted record is treated as "limit reached", not as a reset.
    if (!u) {
      let hasRaw = false;
      try { hasRaw = localStorage.getItem(USAGE_KEY) !== null; } catch (e) {}
      return hasRaw ? CONFIG.FREE_COMPARISONS : 0;
    }
    return Math.max(0, parseInt(u.n, 10) || 0);
  }
  function incrementUsed() { writeStore(USAGE_KEY, { n: getUsed() + 1 }); }
  function freeLeft() { return Math.max(0, CONFIG.FREE_COMPARISONS - getUsed()); }

  /* ---------- Microsoft Store billing (Digital Goods API) ---------- */
  async function getService() {
    if (typeof window === 'undefined' || !('getDigitalGoodsService' in window)) return null;
    try { return await window.getDigitalGoodsService(CONFIG.BILLING_URL); }
    catch (e) { return null; }
  }

  async function refreshEntitlement() {
    const cached = readStore(ENT_KEY);
    const service = await getService();
    if (service) {
      try {
        const purchases = await service.listPurchases();
        const active = (purchases || []).some(p => CONFIG.PRODUCT_IDS.includes(p.itemId));
        premium = active;
        writeStore(ENT_KEY, { ok: active, ts: Date.now() });
        return premium;
      } catch (e) { /* fall through to cached value */ }
    }
    // Store unreachable: honour a recent positive verification (offline grace period)
    const graceMs = CONFIG.OFFLINE_GRACE_DAYS * 86400000;
    premium = !!(cached && cached.ok && (Date.now() - cached.ts) < graceMs);
    return premium;
  }

  async function loadPrice() {
    const service = await getService();
    if (!service) return '';
    try {
      const d = await service.getDetails(CONFIG.PRODUCT_IDS);
      if (d && d[0] && d[0].price) {
        const v = Number(d[0].price.value);
        try { return new Intl.NumberFormat(undefined, { style: 'currency', currency: d[0].price.currency }).format(v); }
        catch (e) { return d[0].price.currency + ' ' + d[0].price.value; }
      }
    } catch (e) {}
    return '';
  }

  async function purchase() {
    const service = await getService();
    if (!service) return { ok: false, reason: 'unavailable' };
    try {
      const details = await service.getDetails(CONFIG.PRODUCT_IDS);
      const item = details && details[0];
      if (!item) return { ok: false, reason: 'noproduct' };
      const request = new PaymentRequest(
        [{ supportedMethods: CONFIG.BILLING_URL, data: { sku: item.itemId } }],
        { total: { label: item.title || CONFIG.APP_NAME, amount: { currency: item.price.currency, value: item.price.value } } }
      );
      const response = await request.show();
      await response.complete('success');
      const active = await refreshEntitlement();
      return { ok: active, reason: active ? 'ok' : 'pending' };
    } catch (e) {
      return { ok: false, reason: (e && e.name === 'AbortError') ? 'cancelled' : 'error' };
    }
  }

  /* ---------- gate ---------- */
  function canCompare() { return !CONFIG.ENABLED || premium || freeLeft() > 0; }

  /* ---------- UI ---------- */
  function injectStyles() {
    const css = `
    .xc-badge{margin-left:auto;font-size:12px;font-weight:600;padding:5px 12px;border-radius:999px;background:rgba(255,255,255,.16);color:#fff;cursor:pointer;border:1px solid rgba(255,255,255,.3)}
    .xc-badge:hover{background:rgba(255,255,255,.26)}
    .xc-overlay{position:fixed;inset:0;background:rgba(20,24,40,.55);display:flex;align-items:center;justify-content:center;z-index:1000;padding:20px}
    .xc-modal{background:#fff;border-radius:14px;max-width:460px;width:100%;padding:28px;box-shadow:0 20px 60px rgba(0,0,0,.3);font-family:var(--sans)}
    .xc-modal h3{margin:0 0 8px;font-size:19px;color:#1B222C}
    .xc-modal p{margin:8px 0;font-size:13.5px;line-height:1.55;color:#5B6472}
    .xc-price{font-size:26px;font-weight:700;color:#2F3B87;margin:14px 0 2px}
    .xc-price small{font-size:13px;font-weight:600;color:#5B6472}
    .xc-actions{display:flex;gap:10px;margin-top:18px;flex-wrap:wrap}
    .xc-msg{font-size:12.5px;margin-top:12px;min-height:16px}
    .xc-msg.err{color:#B23B30}.xc-msg.ok{color:#177A45}
    .xc-fine{font-size:11.5px!important;color:#7A8290!important}
    .xc-link{background:none;border:none;color:#2F3B87;text-decoration:underline;cursor:pointer;font-size:13px;padding:0}`;
    const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
  }

  function updateBadge() {
    const b = document.getElementById('xcBadge');
    if (!b || !CONFIG.ENABLED) { if (b) b.style.display = 'none'; return; }
    b.textContent = premium ? 'Premium active'
      : 'Free comparisons left: ' + freeLeft() + ' of ' + CONFIG.FREE_COMPARISONS;
    b.title = premium ? 'Your subscription is active' : 'Click to see subscription options';
  }

  function setMsg(text, kind) {
    const m = document.getElementById('xcMsg');
    if (m) { m.textContent = text || ''; m.className = 'xc-msg ' + (kind || ''); }
  }

  async function showPaywall(reason) {
    if (document.getElementById('xcOverlay')) return;
    const service = await getService();
    if (service) priceText = await loadPrice();
    const ov = document.createElement('div');
    ov.className = 'xc-overlay'; ov.id = 'xcOverlay';
    const limitHit = freeLeft() === 0;
    const headline = limitHit
      ? 'You\u2019ve used your ' + CONFIG.FREE_COMPARISONS + ' free comparisons'
      : CONFIG.APP_NAME + ' Premium';
    const priceBlock = priceText
      ? '<div class="xc-price">' + priceText + ' <small>/ month</small></div>'
      : '<div class="xc-price"><small>Monthly subscription</small></div>';
    const buyArea = service
      ? '<div class="xc-actions"><button class="primary" id="xcBuy">Subscribe</button>' +
        '<button class="secondary" id="xcRestore">Restore purchase</button>' +
        '<button class="xc-link" id="xcClose">Not now</button></div>'
      : '<p>Subscriptions are available in the Microsoft Store version of ' + CONFIG.APP_NAME + '.</p>' +
        '<div class="xc-actions"><a href="' + CONFIG.STORE_PAGE + '" target="_blank" rel="noopener"><button class="primary">Open in Microsoft Store</button></a>' +
        '<button class="xc-link" id="xcClose">Close</button></div>';
    ov.innerHTML = '<div class="xc-modal" role="dialog" aria-modal="true" aria-labelledby="xcTitle">' +
      '<h3 id="xcTitle">' + headline + '</h3>' +
      '<p>Subscribe to keep comparing workbooks, with unlimited comparisons, Excel and PDF change-control reports, and all features.</p>' +
      priceBlock + buyArea + '<div class="xc-msg" id="xcMsg"></div>' +
      '<p class="xc-fine">The subscription renews automatically every month until you cancel. Payment is charged to your Microsoft account through the Microsoft Store. ' +
      'You can manage or cancel it at any time at account.microsoft.com/services. Your workbooks are never uploaded or shared.</p></div>';
    document.body.appendChild(ov);

    const close = () => { ov.remove(); };
    ov.addEventListener('click', e => { if (e.target === ov) close(); });
    document.getElementById('xcClose').addEventListener('click', close);
    const buy = document.getElementById('xcBuy');
    if (buy) buy.addEventListener('click', async () => {
      if (busy) return; busy = true; setMsg('Opening Microsoft Store checkout\u2026');
      const r = await purchase(); busy = false;
      if (r.ok) { setMsg('Thank you! Premium is now active.', 'ok'); updateBadge(); setTimeout(close, 1200); }
      else if (r.reason === 'cancelled') setMsg('Purchase cancelled.', '');
      else if (r.reason === 'pending') setMsg('Purchase received. It may take a minute to activate; try \u201cRestore purchase\u201d.', '');
      else setMsg('The purchase could not be completed. You have not been charged. Please try again.', 'err');
    });
    const restore = document.getElementById('xcRestore');
    if (restore) restore.addEventListener('click', async () => {
      if (busy) return; busy = true; setMsg('Checking your subscription\u2026');
      const active = await refreshEntitlement(); busy = false; updateBadge();
      if (active) { setMsg('Subscription found. Premium is active.', 'ok'); setTimeout(close, 1200); }
      else setMsg('No active subscription found for this Microsoft account.', 'err');
    });
  }

  /* ---------- wiring ---------- */
  async function gatedCompare(runFn) {
    if (!CONFIG.ENABLED) { runFn(); return; }
    await refreshEntitlement();
    if (!canCompare()) { updateBadge(); showPaywall('limit'); return; }
    const success = runFn();
    if (success !== false && !premium) { incrementUsed(); }
    updateBadge();
  }

  async function init(runFn) {
    if (typeof document === 'undefined') return;
    injectStyles();
    const header = document.querySelector('header');
    if (header) {
      const b = document.createElement('div');
      b.className = 'xc-badge'; b.id = 'xcBadge';
      b.addEventListener('click', () => showPaywall('badge'));
      header.appendChild(b);
    }
    const btn = document.getElementById('btnCompare');
    if (btn) btn.addEventListener('click', () => gatedCompare(runFn));
    updateBadge();
    await refreshEntitlement();
    updateBadge();
  }

  return { init, gatedCompare, refreshEntitlement, purchase, getUsed, incrementUsed, freeLeft, canCompare,
           _config: CONFIG, _isPremium: () => premium };
})();

if (typeof module !== 'undefined') module.exports = Monetization;
