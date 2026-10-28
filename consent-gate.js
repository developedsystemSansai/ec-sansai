/*!
 * consent-gate.js  v2.0  —  EC Sansai
 * ด่านยอมรับ PDPA / คุกกี้ / เงื่อนไขการใช้งาน ก่อนเข้าใช้เว็บไซต์ (ใช้ร่วมทุกหน้า)
 *
 * วิธีใช้ (ใส่ใน <head> ของทุกหน้า หลัง config.js):
 *   <script src="consent-gate.js"></script>
 *   หน้านโยบาย (อ่านได้โดยยังไม่ต้องยอมรับ):  <script src="consent-gate.js" data-consent="soft"></script>
 *   หน้าที่เดิมมี <script data-manual> ของ analytics:  เพิ่ม data-analytics-manual="true"
 *
 * หลักการ
 *  - ไม่มี token ที่ server ออกให้ = ใช้งานไม่ได้: หน้าถูก inert + ทุก request ไป GAS รอจนกว่าจะยอมรับ
 *  - เก็บ token ใน localStorage (ใช้ร่วมทุกหน้าใน origin เดียวกัน) และแนบ "consentToken" ให้ทุก POST ไป GAS_URL อัตโนมัติ
 *    (ครอบ window.fetch) → ไม่ต้องแก้จุดเรียก API ทีละหน้า
 *  - server ตอบ code = CONSENT_REQUIRED (token หมดอายุ/ถูกถอน) → เปิดด่านใหม่แล้วส่งซ้ำให้เอง 1 ครั้ง
 *  - คุกกี้วิเคราะห์ (analytics) เป็นตัวเลือกแยก ค่าเริ่มต้น "ปิด" และโหลด script วิเคราะห์เฉพาะเมื่อผู้ใช้เปิดเท่านั้น
 */
(function () {
  'use strict';
  if (window.ECConsent) return;

  var VERSION = '2.0';
  var LS_KEY = 'ec_consent_v2';
  var LS_DEV = 'ec_device_id';
  var ANALYTICS_SRC = 'https://sandtools.vercel.app/t.js';
  var ANALYTICS_SITE = 'ppucg5v4xb';
  var DPO_TEL = '053-921-199 ต่อ 168';
  var DPO_MAIL = 'sireetorn.wa@sansaihospital.go.th';

  var script = document.currentScript || (function () {
    var s = document.getElementsByTagName('script');
    for (var i = s.length - 1; i >= 0; i--) if (/consent-gate\.js/.test(s[i].src)) return s[i];
    return null;
  })();
  var SOFT = !!(script && script.getAttribute('data-consent') === 'soft');
  var ANALYTICS_MANUAL = !!(script && script.getAttribute('data-analytics-manual') === 'true');

  function gasUrl() {
    if (window.EC_CONFIG && window.EC_CONFIG.GAS_URL) return window.EC_CONFIG.GAS_URL;
    return typeof window.GAS_URL === 'string' ? window.GAS_URL : '';
  }

  /* ───────────── storage (กรณี localStorage ใช้ไม่ได้ เช่น private mode ให้ใช้ memory) ───────────── */
  var mem = {};
  function sGet(k) { try { return localStorage.getItem(k); } catch (e) { return mem[k] || null; } }
  function sSet(k, v) { try { localStorage.setItem(k, v); } catch (e) { mem[k] = v; } }
  function sDel(k) { try { localStorage.removeItem(k); } catch (e) { delete mem[k]; } }

  function uuid() {
    try { if (crypto && crypto.randomUUID) return crypto.randomUUID(); } catch (e) {}
    return 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
  }
  function deviceId() {
    var d = sGet(LS_DEV);
    if (!d) { d = uuid(); sSet(LS_DEV, d); }
    return d;
  }
  function load() {
    try {
      var r = JSON.parse(sGet(LS_KEY) || 'null');
      if (!r || r.version !== VERSION || !r.token) return null;
      if (!r.expiresAt || new Date(r.expiresAt).getTime() < Date.now()) return null;
      return r;
    } catch (e) { return null; }
  }
  function save(rec) { sSet(LS_KEY, JSON.stringify(rec)); }
  function clear() { sDel(LS_KEY); }

  /* ───────────── state ───────────── */
  var state = { rec: load(), open: false, mode: 'gate' };
  var readyResolve;
  var readyPromise = new Promise(function (res) { readyResolve = res; });
  if (state.rec) readyResolve(state.rec);

  /* ───────────── analytics (โหลดเฉพาะเมื่อผู้ใช้เปิดไว้) ───────────── */
  var analyticsLoaded = false;
  function applyAnalytics() {
    if (analyticsLoaded || !state.rec || !state.rec.analytics) return;
    analyticsLoaded = true;
    var s = document.createElement('script');
    s.defer = true; s.src = ANALYTICS_SRC; s.setAttribute('data-site', ANALYTICS_SITE);
    if (ANALYTICS_MANUAL) s.setAttribute('data-manual', '');
    document.head.appendChild(s);
  }

  /* ───────────── fetch wrapper: รอ consent → แนบ token → ส่งซ้ำเมื่อ server สั่งให้ยอมรับใหม่ ───────────── */
  var origFetch = window.fetch ? window.fetch.bind(window) : null;
  var EXEMPT = { recordPdpaConsent: 1, revokePdpaConsent: 1 };

  function isGasRequest(input, init) {
    var url = typeof input === 'string' ? input : (input && input.url) || '';
    var g = gasUrl();
    return !!(g && url && url.indexOf(g) === 0 && init && typeof init.body === 'string' &&
              String((init.method || 'GET')).toUpperCase() === 'POST');
  }
  function withToken(init, token) {
    var body;
    try { body = JSON.parse(init.body); } catch (e) { return null; }
    if (!body || typeof body !== 'object' || EXEMPT[body.action]) return { init: init, exempt: true };
    body.consentToken = token;
    var copy = {}; for (var k in init) copy[k] = init[k];
    copy.body = JSON.stringify(body);
    return { init: copy, exempt: false };
  }

  if (origFetch) {
    window.fetch = function (input, init) {
      if (!isGasRequest(input, init)) return origFetch(input, init);
      var action = '';
      try { action = JSON.parse(init.body).action; } catch (e) {}
      if (EXEMPT[action]) return origFetch(input, init);

      function attempt(retried) {
        return readyPromise.then(function (rec) {
          var w = withToken(init, rec.token);
          if (!w) return origFetch(input, init);
          return origFetch(input, w.init).then(function (resp) {
            if (retried) return resp;
            return resp.clone().text().then(function (t) {
              var s = t.indexOf('{'), e = t.lastIndexOf('}'), j = null;
              if (s !== -1 && e > s) { try { j = JSON.parse(t.substring(s, e + 1)); } catch (x) {} }
              if (j && j.code === 'CONSENT_REQUIRED') {
                resetReady();
                openGate('gate');
                return attempt(true);
              }
              return resp;
            }).catch(function () { return resp; });
          });
        });
      }
      return attempt(false);
    };
  }

  function resetReady() {
    clear(); state.rec = null;
    readyPromise = new Promise(function (res) { readyResolve = res; });
  }

  /* ───────────── เรียก server (ไม่ผ่านตัวครอบ) ───────────── */
  function post(action, payload) {
    var g = gasUrl();
    if (!g || !origFetch) return Promise.reject(new Error('no-endpoint'));
    return origFetch(g, {
      method: 'POST', headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action: action, origin: window.location.origin, payload: payload })
    }).then(function (r) { return r.text(); }).then(function (t) {
      var s = t.indexOf('{'), e = t.lastIndexOf('}');
      if (s === -1 || e <= s) throw new Error('bad-response');
      return JSON.parse(t.substring(s, e + 1));
    });
  }
  // IP ใช้เป็นหลักฐานการยินยอมเท่านั้น และเรียกหลังผู้ใช้กดยอมรับแล้ว (ไม่เรียกบุคคลที่สามก่อนได้รับความยินยอม)
  function getIp(cb) {
    var done = false;
    function fin(v) { if (!done) { done = true; cb(v || ''); } }
    try {
      var c = new AbortController(); var t = setTimeout(function () { c.abort(); fin(''); }, 2500);
      origFetch('https://api.ipify.org?format=json', { signal: c.signal })
        .then(function (r) { return r.json(); })
        .then(function (j) { clearTimeout(t); fin(j && j.ip); })
        .catch(function () { clearTimeout(t); fin(''); });
    } catch (e) { fin(''); }
  }

  /* ───────────── UI ───────────── */
  var CSS = [
    '#ecg-root,#ecg-root *{box-sizing:border-box}',
    '#ecg-root{position:fixed;inset:0;z-index:2147483000;display:flex;align-items:center;justify-content:center;padding:16px;',
    '  background:rgba(11,31,53,.78);-webkit-backdrop-filter:blur(3px);backdrop-filter:blur(3px);font-family:Sarabun,"Helvetica Neue",Arial,sans-serif;color:#151e28}',
    '#ecg-root[hidden]{display:none}',
    '.ecg-card{background:#fff;border-radius:16px;box-shadow:0 24px 70px rgba(11,31,53,.45);width:100%;max-width:640px;max-height:calc(100vh - 32px);display:flex;flex-direction:column;overflow:hidden;border-top:4px solid #b58a2a}',
    '.ecg-hd{display:flex;gap:14px;align-items:center;padding:20px 24px 12px}',
    '.ecg-ico{flex:0 0 46px;height:46px;border-radius:12px;background:#0b1f35;color:#c9a037;display:grid;place-items:center}',
    '.ecg-hd h2{margin:0;font-family:"Noto Serif Thai",Georgia,serif;font-size:1.15rem;line-height:1.35;color:#0b1f35}',
    '.ecg-hd p{margin:2px 0 0;font-size:.82rem;color:#5a6e85}',
    '.ecg-bd{padding:4px 24px 8px;overflow:auto;flex:1}',
    '.ecg-lead{font-size:.92rem;line-height:1.75;color:#344255;margin:0 0 12px}',
    '.ecg-sec{border:1px solid #d0dae6;border-radius:12px;margin:0 0 10px;background:#f4f6fa}',
    '.ecg-sec>summary{cursor:pointer;list-style:none;padding:12px 14px;font-weight:600;font-size:.92rem;color:#0b1f35;display:flex;align-items:center;gap:10px}',
    '.ecg-sec>summary::-webkit-details-marker{display:none}',
    '.ecg-sec>summary::after{content:"";margin-left:auto;width:8px;height:8px;border-right:2px solid #5a6e85;border-bottom:2px solid #5a6e85;transform:rotate(45deg);transition:transform .2s}',
    '.ecg-sec[open]>summary::after{transform:rotate(-135deg)}',
    '.ecg-sec .ecg-num{flex:0 0 24px;height:24px;border-radius:50%;background:#0b1f35;color:#fff;font-size:.75rem;display:grid;place-items:center}',
    '.ecg-sec .ecg-body{padding:0 14px 14px 48px;font-size:.84rem;line-height:1.75;color:#344255}',
    '.ecg-body ul{margin:6px 0 6px 1.1em;padding:0}.ecg-body li{margin:2px 0}',
    '.ecg-body a{color:#9a7220;font-weight:600}',
    '.ecg-row{display:flex;align-items:center;justify-content:space-between;gap:12px;background:#fff;border:1px solid #e8edf5;border-radius:10px;padding:10px 12px;margin-top:8px}',
    '.ecg-row b{display:block;font-size:.86rem;color:#0b1f35}.ecg-row small{color:#5a6e85}',
    '.ecg-tag{font-size:.65rem;font-weight:700;background:#173454;color:#fff;border-radius:20px;padding:2px 8px;margin-left:6px;vertical-align:middle}',
    '.ecg-sw{position:relative;flex:0 0 auto}.ecg-sw input{position:absolute;opacity:0;width:100%;height:100%;margin:0;cursor:pointer}',
    '.ecg-sw i{display:block;width:44px;height:24px;border-radius:24px;background:#a8b8cb;position:relative;transition:background .18s}',
    '.ecg-sw i::after{content:"";position:absolute;top:3px;left:3px;width:18px;height:18px;border-radius:50%;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.3);transition:transform .18s}',
    '.ecg-sw input:checked+i{background:#1b5e3e}.ecg-sw input:checked+i::after{transform:translateX(20px)}',
    '.ecg-sw input:disabled+i{opacity:.7;cursor:not-allowed}.ecg-sw input:focus-visible+i{outline:2px solid #b58a2a;outline-offset:2px}',
    '.ecg-ft{padding:14px 24px 20px;border-top:1px solid #e8edf5;background:#fff}',
    '.ecg-chk{display:flex;gap:10px;align-items:flex-start;font-size:.88rem;line-height:1.6;color:#151e28;cursor:pointer;margin:0 0 12px}',
    '.ecg-chk input{margin-top:4px;width:18px;height:18px;accent-color:#1b5e3e;flex:0 0 auto}',
    '.ecg-act{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}',
    '.ecg-btn{font:600 .92rem Sarabun,Arial,sans-serif;border-radius:8px;padding:11px 20px;min-height:44px;cursor:pointer;border:1.5px solid transparent}',
    '.ecg-btn:focus-visible{outline:3px solid #b58a2a;outline-offset:2px}',
    '.ecg-pri{background:#b58a2a;color:#0b1f35}.ecg-pri:hover:not(:disabled){background:#c9a037}',
    '.ecg-pri:disabled{background:#d0dae6;color:#5a6e85;cursor:not-allowed}',
    '.ecg-sec2{background:#fff;border-color:#d0dae6;color:#344255}.ecg-sec2:hover{background:#f4f6fa}',
    '.ecg-danger{background:#fff;border-color:#b91c1c;color:#b91c1c}.ecg-danger:hover{background:#fef2f2}',
    '.ecg-err{display:none;background:#fef2f2;border:1px solid #fecaca;color:#b91c1c;border-radius:8px;padding:8px 12px;font-size:.84rem;margin:0 0 10px}',
    '.ecg-err.on{display:block}',
    '.ecg-decl{padding:8px 24px 24px;text-align:center}.ecg-decl p{color:#344255;line-height:1.8}',
    '#ecg-pill{position:fixed;left:12px;bottom:12px;z-index:9990;background:#0b1f35;color:#fff;border:1px solid #b58a2a;border-radius:999px;',
    '  font:600 .78rem Sarabun,Arial,sans-serif;padding:8px 14px;cursor:pointer;display:none;align-items:center;gap:6px;box-shadow:0 4px 14px rgba(11,31,53,.35)}',
    '#ecg-pill.on{display:inline-flex}#ecg-pill:hover{background:#173454}#ecg-pill:focus-visible{outline:3px solid #b58a2a;outline-offset:2px}',
    '#ecg-soft{position:fixed;left:0;right:0;bottom:0;z-index:2147482000;background:#0b1f35;color:#fff;padding:12px 16px;display:flex;gap:12px;align-items:center;justify-content:center;flex-wrap:wrap;font:.88rem Sarabun,Arial,sans-serif;border-top:3px solid #b58a2a}',
    '#ecg-soft button{background:#b58a2a;border:0;color:#0b1f35;font:700 .88rem Sarabun,Arial,sans-serif;border-radius:8px;padding:9px 18px;cursor:pointer}',
    'html.ecg-lock,html.ecg-lock body{overflow:hidden!important}',
    '@media(max-width:560px){.ecg-hd{padding:16px 16px 8px}.ecg-bd{padding:4px 16px 8px}.ecg-ft{padding:12px 16px 16px}.ecg-sec .ecg-body{padding-left:14px}.ecg-act .ecg-btn{flex:1 1 100%}}',
    '@media(prefers-reduced-motion:reduce){*{transition:none!important}}'
  ].join('\n');

  var ICON_SHIELD = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6l8-3z"/><path d="M9 12l2 2 4-4"/></svg>';

  var root, pill, softBar, lastFocus, inertTouched = [];

  function el(tag, attrs, html) {
    var n = document.createElement(tag);
    if (attrs) for (var k in attrs) n.setAttribute(k, attrs[k]);
    if (html != null) n.innerHTML = html;
    return n;
  }
  function injectCss() {
    if (document.getElementById('ecg-css')) return;
    var s = el('style', { id: 'ecg-css' }); s.textContent = CSS; document.head.appendChild(s);
  }

  function gateHtml(analyticsOn, settings) {
    var title = settings ? 'ตั้งค่าความเป็นส่วนตัวและความยินยอม' : 'ก่อนเข้าใช้งานเว็บไซต์';
    var sub = settings ? 'Privacy & consent settings' : 'Please review and accept to continue';
    var lead = settings
      ? 'ท่านให้ความยินยอมไว้แล้ว สามารถเปลี่ยนตัวเลือกคุกกี้วิเคราะห์ หรือถอนความยินยอมได้ตลอดเวลา การถอนความยินยอมจะทำให้ไม่สามารถใช้งานเว็บไซต์ได้จนกว่าจะยอมรับใหม่'
      : 'เว็บไซต์ระบบยื่นจริยธรรมการวิจัยในมนุษย์ โรงพยาบาลสันทราย ต้องได้รับการยอมรับนโยบายความเป็นส่วนตัว (PDPA) คุกกี้ และเงื่อนไขการใช้งานด้านล่าง ก่อนจึงจะเข้าใช้งานได้';
    return '' +
    '<div class="ecg-card" role="dialog" aria-modal="true" aria-labelledby="ecg-title" aria-describedby="ecg-lead">' +
      '<div class="ecg-hd"><div class="ecg-ico">' + ICON_SHIELD + '</div><div><h2 id="ecg-title" tabindex="-1">' + title + '</h2><p>' + sub + '</p></div></div>' +
      '<div class="ecg-bd">' +
        '<p class="ecg-lead" id="ecg-lead">' + lead + '</p>' +
        '<div id="ecg-main">' +
        '<details class="ecg-sec" open><summary><span class="ecg-num">1</span>นโยบายความเป็นส่วนตัว (PDPA)</summary><div class="ecg-body">' +
          '<p>โรงพยาบาลสันทรายเก็บและใช้ข้อมูลส่วนบุคคล เช่น ชื่อ อีเมล เบอร์โทร สังกัด และข้อมูลโครงการวิจัย เพื่อบริหารงานพิจารณาจริยธรรมการวิจัยในมนุษย์เท่านั้น ตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. 2562</p>' +
          '<ul><li>เมื่อท่านกดยอมรับ ระบบจะบันทึกเวลา เวอร์ชันของนโยบาย ตัวเลือกที่ท่านเลือก ประเภทเบราว์เซอร์ และหมายเลข IP เพื่อเป็นหลักฐานการให้ความยินยอม</li>' +
          '<li>ท่านมีสิทธิขอเข้าถึง แก้ไข ลบ หรือคัดค้านการใช้ข้อมูล ติดต่อเจ้าหน้าที่คุ้มครองข้อมูล โทร ' + DPO_TEL + ' หรืออีเมล <a href="mailto:' + DPO_MAIL + '">' + DPO_MAIL + '</a></li></ul>' +
          '<p><a href="privacy-policy.html" target="_blank" rel="noopener">อ่านนโยบายความเป็นส่วนตัวฉบับเต็ม (เปิดแท็บใหม่)</a></p>' +
        '</div></details>' +
        '<details class="ecg-sec"><summary><span class="ecg-num">2</span>คุกกี้</summary><div class="ecg-body">' +
          '<div class="ecg-row"><div><b>คุกกี้ที่จำเป็น <span class="ecg-tag">จำเป็น</span></b><small>การเข้าสู่ระบบ ความปลอดภัย และจดจำการให้ความยินยอมนี้</small></div>' +
            '<label class="ecg-sw"><input type="checkbox" checked disabled aria-label="คุกกี้ที่จำเป็น (ปิดไม่ได้)"><i></i></label></div>' +
          '<div class="ecg-row"><div><b>คุกกี้เพื่อการวิเคราะห์</b><small>นับจำนวนผู้เข้าชมเพื่อปรับปรุงเว็บไซต์ — ไม่บังคับ ค่าเริ่มต้นปิด</small></div>' +
            '<label class="ecg-sw"><input type="checkbox" id="ecg-analytics" ' + (analyticsOn ? 'checked' : '') + ' aria-label="คุกกี้เพื่อการวิเคราะห์ (ไม่บังคับ)"><i></i></label></div>' +
        '</div></details>' +
        '<details class="ecg-sec"><summary><span class="ecg-num">3</span>เงื่อนไขการเผยแพร่ผลงานวิจัยหลังปิดโครงการ</summary><div class="ecg-body">' +
          '<p>เมื่อผู้วิจัยดำเนินการ <b>ปิดโครงการ</b> เรียบร้อยแล้ว สำนักงานวิจัยในมนุษย์ โรงพยาบาลสันทราย ขออนุญาตนำผลการวิจัยฉบับย่อไปจัดทำเป็น <b>Infographic สั้น กระชับ</b> เพื่อเผยแพร่บนเว็บไซต์ เป็นผลงานของโรงพยาบาล โดย</p>' +
          '<ul><li>เผยแพร่เฉพาะสรุปผลโดยย่อ ไม่เผยแพร่ข้อมูลที่ระบุตัวตนของอาสาสมัครหรือผู้เข้าร่วมวิจัย</li>' +
          '<li>ชื่อผู้วิจัยจะแสดงต่อเมื่อท่านให้ความยินยอมเพิ่มเติมแยกต่างหากตอนยื่นรายงานปิดโครงการเท่านั้น</li>' +
          '<li>ลิขสิทธิ์และความเป็นเจ้าของผลงานวิจัยยังเป็นของผู้วิจัยตามเดิม</li>' +
          '<li>ท่านขอแก้ไขหรือถอดผลงานออกได้ทุกเมื่อ โดยติดต่อสำนักงานวิจัยในมนุษย์ โทร ' + DPO_TEL + '</li></ul>' +
        '</div></details>' +
        '</div>' +
        '<div class="ecg-decl" id="ecg-decl" hidden><h3 style="font-family:\'Noto Serif Thai\',serif;margin:8px 0">ไม่สามารถใช้งานเว็บไซต์ได้</h3>' +
          '<p>การใช้งานระบบต้องได้รับความยินยอมตามข้อ 1–3 ท่านยังสามารถอ่านนโยบายฉบับเต็ม หรือติดต่อเจ้าหน้าที่ได้ที่ ' + DPO_TEL + '</p>' +
          '<div class="ecg-act" style="justify-content:center"><button class="ecg-btn ecg-sec2" id="ecg-back" type="button">กลับไปอ่านเงื่อนไข</button>' +
          '<a class="ecg-btn ecg-sec2" href="privacy-policy.html" style="text-decoration:none;display:inline-grid;place-items:center" target="_blank" rel="noopener">อ่านนโยบายฉบับเต็ม</a></div></div>' +
      '</div>' +
      '<div class="ecg-ft" id="ecg-ft">' +
        '<div class="ecg-err" id="ecg-err" role="alert"></div>' +
        '<label class="ecg-chk"><input type="checkbox" id="ecg-agree"' + (settings ? ' checked' : '') + '><span>ข้าพเจ้าได้อ่านและยอมรับ <b>นโยบายความเป็นส่วนตัว (PDPA)</b> และ <b>เงื่อนไขการใช้งาน</b> ข้างต้น รวมถึงเงื่อนไขการเผยแพร่ผลงานวิจัยหลังปิดโครงการ</span></label>' +
        '<div class="ecg-act">' +
          (settings ? '<button class="ecg-btn ecg-danger" id="ecg-withdraw" type="button">ถอนความยินยอม</button><button class="ecg-btn ecg-sec2" id="ecg-close" type="button">ปิด</button>'
                    : (SOFT ? '<button class="ecg-btn ecg-sec2" id="ecg-close" type="button">กลับไปอ่านนโยบายต่อ</button>' : '<button class="ecg-btn ecg-sec2" id="ecg-decline" type="button">ไม่ยอมรับ</button>')) +
          '<button class="ecg-btn ecg-pri" id="ecg-accept" type="button"' + (settings ? '' : ' disabled') + '>' + (settings ? 'บันทึกการตั้งค่า' : 'ยอมรับและเข้าใช้งาน') + '</button>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  /* inert ทั้งหน้ายกเว้นด่าน → โฟกัส/คลิก/screen reader เข้าถึงเนื้อหาด้านหลังไม่ได้ */
  function lockPage(on) {
    var docEl = document.documentElement;
    if (on) {
      docEl.classList.add('ecg-lock');
      inertTouched = [];
      Array.prototype.forEach.call(document.body.children, function (n) {
        if (n === root || n === pill || n === softBar || n.tagName === 'SCRIPT' || n.tagName === 'STYLE') return;
        if (!n.hasAttribute('inert')) { n.setAttribute('inert', ''); n.setAttribute('data-ecg-inert', '1'); inertTouched.push(n); }
      });
    } else {
      docEl.classList.remove('ecg-lock');
      Array.prototype.forEach.call(document.querySelectorAll('[data-ecg-inert]'), function (n) {
        n.removeAttribute('inert'); n.removeAttribute('data-ecg-inert');
      });
      inertTouched = [];
    }
  }

  function trapKeys(e) {
    if (!state.open || !root || root.hidden) return;
    if (e.key === 'Escape') {
      if (state.mode === 'settings' || SOFT) { e.preventDefault(); closeGate(); }
      return; // โหมดบังคับ: Esc ปิดไม่ได้
    }
    if (e.key !== 'Tab') return;
    var f = root.querySelectorAll('button:not([disabled]),a[href],input:not([disabled]),summary,[tabindex="0"]');
    if (!f.length) return;
    var first = f[0], last = f[f.length - 1];
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  }

  function showErr(msg) {
    var b = root && root.querySelector('#ecg-err'); if (!b) return;
    b.textContent = msg || ''; b.classList.toggle('on', !!msg);
  }

  function openGate(mode) {
    if (!document.body) { document.addEventListener('DOMContentLoaded', function () { openGate(mode); }); return; }
    injectCss();
    state.mode = mode; state.open = true;
    lastFocus = document.activeElement;
    if (!root) {
      root = el('div', { id: 'ecg-root' });
      document.body.appendChild(root);
      document.addEventListener('keydown', trapKeys, true);
    }
    var rec = state.rec || {};
    root.innerHTML = gateHtml(!!rec.analytics, mode === 'settings');
    root.hidden = false;
    if (softBar) softBar.style.display = 'none';
    if (pill) pill.classList.remove('on');
    lockPage(true);
    bind();
    var t = root.querySelector('#ecg-title'); if (t) t.focus();
  }

  function closeGate() {
    state.open = false;
    if (root) root.hidden = true;
    lockPage(false);
    if (state.rec) showPill(); else if (SOFT) showSoft();
    try { if (lastFocus && lastFocus.focus) lastFocus.focus(); } catch (e) {}
  }

  function bind() {
    var agree = root.querySelector('#ecg-agree'), accept = root.querySelector('#ecg-accept');
    var analytics = root.querySelector('#ecg-analytics');
    var decline = root.querySelector('#ecg-decline'), back = root.querySelector('#ecg-back');
    var closeB = root.querySelector('#ecg-close'), withdraw = root.querySelector('#ecg-withdraw');

    if (agree && state.mode !== 'settings') agree.addEventListener('change', function () { accept.disabled = !agree.checked; });
    if (agree && state.mode === 'settings') agree.disabled = true; // ตั้งค่า: ยินยอมอยู่แล้ว (ถอนผ่านปุ่มถอนเท่านั้น)
    if (decline) decline.addEventListener('click', function () {
      root.querySelector('#ecg-main').hidden = true; root.querySelector('#ecg-ft').hidden = true; root.querySelector('#ecg-lead').hidden = true;
      root.querySelector('#ecg-decl').hidden = false; var b = root.querySelector('#ecg-back'); if (b) b.focus();
    });
    if (back) back.addEventListener('click', function () {
      root.querySelector('#ecg-main').hidden = false; root.querySelector('#ecg-ft').hidden = false; root.querySelector('#ecg-lead').hidden = false;
      root.querySelector('#ecg-decl').hidden = true; if (agree) agree.focus();
    });
    if (closeB) closeB.addEventListener('click', closeGate);

    if (accept) accept.addEventListener('click', function () {
      if (state.mode !== 'settings' && !(agree && agree.checked)) return;
      accept.disabled = true; var label = accept.textContent; accept.textContent = 'กำลังบันทึก…'; showErr('');
      getIp(function (ip) {
        post('recordPdpaConsent', {
          version: VERSION, acceptNotice: true, publicationAck: true,
          analytics: !!(analytics && analytics.checked),
          deviceId: deviceId(), ip: ip, userAgent: navigator.userAgent,
          page: (location.pathname.split('/').pop() || 'index.html')
        }).then(function (res) {
          if (!res || !res.success || !res.token) {
            accept.disabled = false; accept.textContent = label;
            showErr((res && res.message) || 'ไม่สามารถบันทึกการยินยอมได้ กรุณาลองใหม่');
            return;
          }
          state.rec = { version: VERSION, token: res.token, consentId: res.consentId, expiresAt: res.expiresAt,
                        analytics: !!(analytics && analytics.checked), savedAt: new Date().toISOString() };
          save(state.rec);
          readyResolve(state.rec);
          closeGate(); applyAnalytics();
          try { window.dispatchEvent(new CustomEvent('ec:consent', { detail: { analytics: state.rec.analytics } })); } catch (e) {}
        }).catch(function () {
          accept.disabled = false; accept.textContent = label;
          showErr('เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ กรุณาตรวจสอบอินเทอร์เน็ตแล้วลองอีกครั้ง');
        });
      });
    });

    if (withdraw) withdraw.addEventListener('click', function () {
      if (!window.confirm('ยืนยันการถอนความยินยอม? ท่านจะใช้งานเว็บไซต์ไม่ได้จนกว่าจะยอมรับใหม่')) return;
      withdraw.disabled = true;
      var tok = state.rec && state.rec.token;
      getIp(function (ip) {
        post('revokePdpaConsent', { token: tok, ip: ip, userAgent: navigator.userAgent, page: (location.pathname.split('/').pop() || 'index.html') })
          .catch(function () {})
          .then(function () { resetReady(); window.location.reload(); });
      });
    });
  }

  function showPill() {
    if (SOFT) return;
    if (!pill) {
      pill = el('button', { id: 'ecg-pill', type: 'button', 'aria-label': 'ตั้งค่าความเป็นส่วนตัวและความยินยอม' },
        ICON_SHIELD.replace('width="24" height="24"', 'width="14" height="14"') + '<span>ความเป็นส่วนตัว</span>');
      pill.addEventListener('click', function () { openGate('settings'); });
      document.body.appendChild(pill);
    }
    var cf = document.getElementById('copyright-footer'); // app.html มีแถบลิขสิทธิ์ติดขอบล่าง
    pill.style.bottom = cf ? (cf.offsetHeight + 12) + 'px' : '12px';
    pill.classList.add('on');
  }
  function showSoft() {
    if (!softBar) {
      softBar = el('div', { id: 'ecg-soft', role: 'region', 'aria-label': 'การยอมรับนโยบาย' },
        '<span>ท่านยังไม่ได้ยอมรับนโยบายและเงื่อนไข — อ่านได้ตามสบาย เมื่อพร้อมแล้วกดยอมรับเพื่อใช้งานเว็บไซต์</span><button type="button">ยอมรับและเข้าใช้งาน</button>');
      softBar.querySelector('button').addEventListener('click', function () { openGate('gate'); });
      document.body.appendChild(softBar);
    }
    softBar.style.display = 'flex';
  }

  /* ───────────── public API ───────────── */
  window.ECConsent = {
    version: VERSION,
    ready: function () { return readyPromise; },
    isGranted: function () { return !!state.rec; },
    token: function () { return state.rec ? state.rec.token : ''; },
    analytics: function () { return !!(state.rec && state.rec.analytics); },
    open: function () { openGate(state.rec ? 'settings' : 'gate'); },
    withdraw: function () { openGate('settings'); }
  };

  function start() {
    injectCss();
    if (state.rec) { showPill(); applyAnalytics(); return; }
    if (SOFT) { showSoft(); return; }
    openGate('gate');
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
