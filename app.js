/* TaxiF · Omra 2027 prize draw
 * Everything runs in the browser. Files are never uploaded anywhere.
 */
(() => {
  'use strict';

  const $ = (s, r = document) => r.querySelector(s);
  const TAU = Math.PI * 2;
  const mod = (a, n) => ((a % n) + n) % n;
  const fmt = n => Number(n).toLocaleString('en-US');
  const SRC_COL = 'الملف المصدر';
  const EMPTY_LABEL = '(فارغ)';

  /* ------------------------------------------------------------------ *
   *  Helpers
   * ------------------------------------------------------------------ */
  const norm = s => String(s ?? '')
    .replace(/[ً-ٰٟـ]/g, '')
    .replace(/[إأآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه')
    .toLowerCase().trim();

  function randInt(n) {
    if (n <= 1) return 0;
    const lim = Math.floor(0x100000000 / n) * n;
    const a = new Uint32Array(1);
    do { crypto.getRandomValues(a); } while (a[0] >= lim);
    return a[0] % n;
  }
  function shuffle(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = randInt(i + 1);
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
  }
  function sample(list, n) {
    if (list.length <= n) return shuffle(list.slice());
    const picked = new Set();
    while (picked.size < n) picked.add(randInt(list.length));
    return shuffle([...picked].map(i => list[i]));
  }

  let toastTimer;
  function toast(msg, isErr) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.toggle('err', !!isErr);
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 4200);
  }

  /* ------------------------------------------------------------------ *
   *  Sound (WebAudio, no assets)
   * ------------------------------------------------------------------ */
  const sound = (() => {
    let ac = null, muted = false, lastTick = 0;
    const ctx = () => {
      if (!ac) { const C = window.AudioContext || window.webkitAudioContext; if (C) ac = new C(); }
      if (ac && ac.state === 'suspended') ac.resume();
      return ac;
    };
    function beep(freq, start, dur, type = 'sine', vol = .15) {
      const a = ctx(); if (!a) return;
      const o = a.createOscillator(), g = a.createGain();
      o.type = type; o.frequency.value = freq;
      g.gain.setValueAtTime(0.0001, a.currentTime + start);
      g.gain.exponentialRampToValueAtTime(vol, a.currentTime + start + .02);
      g.gain.exponentialRampToValueAtTime(0.0001, a.currentTime + start + dur);
      o.connect(g).connect(a.destination);
      o.start(a.currentTime + start); o.stop(a.currentTime + start + dur + .05);
    }
    return {
      unlock() { ctx(); },
      get muted() { return muted; },
      set muted(v) { muted = v; },
      tick() {
        if (muted) return;
        const now = performance.now();
        if (now - lastTick < 40) return;
        lastTick = now;
        beep(820 + Math.random() * 120, 0, .05, 'triangle', .07);
      },
      win() {
        if (muted) return;
        [523.25, 659.25, 783.99, 1046.5].forEach((f, i) => beep(f, i * .14, .5, 'triangle', .16));
        [1046.5, 1318.5, 1568].forEach((f, i) => beep(f, .62 + i * .1, .9, 'sine', .12));
      },
    };
  })();

  /* ------------------------------------------------------------------ *
   *  File parsing
   * ------------------------------------------------------------------ */
  function flatten(obj, prefix = '', out = {}, depth = 0) {
    for (const [k, v] of Object.entries(obj)) {
      const key = prefix + k;
      if (v && typeof v === 'object' && !Array.isArray(v) && depth < 3) {
        flatten(v, key + '.', out, depth + 1);
      } else if (Array.isArray(v)) {
        out[key] = v.every(x => x === null || typeof x !== 'object') ? v.join('، ') : JSON.stringify(v);
      } else {
        out[key] = v == null ? '' : String(v);
      }
    }
    return out;
  }

  function pickArray(data, depth = 0) {
    if (Array.isArray(data)) return data;
    if (!data || typeof data !== 'object') return [data];
    const entries = Object.entries(data);
    const arrays = entries.filter(([, v]) => Array.isArray(v)).sort((a, b) => b[1].length - a[1].length);
    if (arrays.length && arrays[0][1].length) return arrays[0][1];
    const objs = entries.filter(([, v]) => v && typeof v === 'object' && !Array.isArray(v));
    if (entries.length > 1 && objs.length === entries.length) return entries.map(([k, v]) => ({ _key: k, ...v })); // keyed by id
    if (objs.length === 1 && depth < 3) return pickArray(objs[0][1], depth + 1);                                   // wrapper
    return [data];
  }

  function jsonToRows(data) {
    return pickArray(data).map(item =>
      item && typeof item === 'object' && !Array.isArray(item) ? flatten(item) : { 'القيمة': item == null ? '' : String(item) });
  }

  function columnsOf(rows) {
    const set = new Set();
    for (const r of rows) for (const k of Object.keys(r)) set.add(k);
    return [...set];
  }

  async function readFile(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    let rows;
    if (ext === 'json') {
      let data;
      try { data = JSON.parse((await file.text()).replace(/^﻿/, '')); }
      catch { throw new Error('ملف JSON غير صالح: ' + file.name); }
      rows = jsonToRows(data);
    } else if (['xlsx', 'xls', 'xlsm', 'ods', 'csv', 'tsv'].includes(ext)) {
      if (!window.XLSX) throw new Error('مكتبة Excel لم تُحمَّل');
      const wb = ['csv', 'tsv'].includes(ext)
        ? XLSX.read(await file.text(), { type: 'string' })
        : XLSX.read(await file.arrayBuffer(), { type: 'array' });
      let best = [];
      for (const name of wb.SheetNames) {
        const r = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: false });
        if (r.length > best.length) best = r;
      }
      rows = best.map(r => Object.fromEntries(Object.entries(r).map(([k, v]) => [String(k).trim(), String(v ?? '').trim()])));
    } else {
      throw new Error('نوع الملف غير مدعوم: ' + file.name);
    }
    rows = rows.filter(r => Object.values(r).some(v => String(v).trim() !== ''));
    if (!rows.length) throw new Error('الملف فارغ: ' + file.name);
    return { name: file.name, ext: ext === 'json' ? 'json' : 'xl', rows, columns: columnsOf(rows) };
  }

  /* ------------------------------------------------------------------ *
   *  Merge two files (auto-detect shared key such as phone / id)
   * ------------------------------------------------------------------ */
  const keyRe = /phone|mobile|tel|id\b|uid|هاتف|جوال|موبايل|رقم|هويه|هوية|معرف|license|رخص/i;

  function normKey(v) {
    let s = String(v ?? '').trim();
    if (!s) return '';
    s = s.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
    if (/^[\d\s+\-().]+$/.test(s)) {
      s = s.replace(/\D/g, '');
      return s.length > 9 ? s.slice(-9) : s.replace(/^0+/, '');
    }
    return norm(s).replace(/\s+/g, ' ');
  }

  function keyColumns(rows, cols) {
    const out = [];
    for (const c of cols) {
      const set = new Set(); let non = 0;
      for (const r of rows) { const k = normKey(r[c]); if (k) { non++; set.add(k); } }
      if (non < rows.length * .5 || set.size < 3 || set.size / non < .9) continue;
      out.push({ c, set });
    }
    return out;
  }

  function detectJoin(rowsA, colsA, rowsB, colsB) {
    const A = keyColumns(rowsA, colsA), B = keyColumns(rowsB, colsB);
    let best = null;
    for (const a of A) for (const b of B) {
      const [small, big] = a.set.size < b.set.size ? [a.set, b.set] : [b.set, a.set];
      let hit = 0;
      for (const k of small) if (big.has(k)) hit++;
      const score = hit / small.size;
      const total = score + (a.c === b.c ? .05 : 0) + (keyRe.test(a.c) ? .02 : 0);
      if (hit >= 2 && score >= .2 && (!best || total > best.total)) best = { a: a.c, b: b.c, hit, total };
    }
    return best;
  }

  function mergeAll(datasets) {
    if (!datasets.length) return { rows: [], cols: [], joins: [] };
    const cols = [...datasets[0].columns];
    const rows = datasets[0].rows.map(d => ({ data: { ...d }, src: [0] }));
    const joins = [];
    for (let i = 1; i < datasets.length; i++) {
      const ds = datasets[i];
      const j = detectJoin(rows.map(r => r.data), cols, ds.rows, ds.columns);
      for (const c of ds.columns) if (!cols.includes(c)) cols.push(c);
      if (j) {
        const idx = new Map();
        for (const r of rows) { const k = normKey(r.data[j.a]); if (k && !idx.has(k)) idx.set(k, r); }
        let matched = 0;
        for (const d of ds.rows) {
          const k = normKey(d[j.b]);
          const t = k && idx.get(k);
          if (t) {
            matched++;
            for (const c of ds.columns) if (t.data[c] === undefined || t.data[c] === '') t.data[c] = d[c];
            if (!t.src.includes(i)) t.src.push(i);
          } else {
            const nd = { ...d };
            if (j.a !== j.b && !nd[j.a]) nd[j.a] = d[j.b];
            rows.push({ data: nd, src: [i] });
          }
        }
        joins.push({ ...j, ds: i, matched });
      } else {
        for (const d of ds.rows) rows.push({ data: { ...d }, src: [i] });
        joins.push({ ds: i, none: true });
      }
    }
    if (datasets.length > 1) {
      cols.push(SRC_COL);
      for (const r of rows) {
        r.data[SRC_COL] = r.src.length === datasets.length
          ? (datasets.length === 2 ? 'كلا الملفين' : 'كل الملفات')
          : r.src.map(i => datasets[i].name).join(' + ');
      }
    }
    rows.forEach((r, i) => { r.id = i; });
    return { rows, cols, joins };
  }

  /* ------------------------------------------------------------------ *
   *  State
   * ------------------------------------------------------------------ */
  const state = {
    datasets: [], rows: [], cols: [], joins: [],
    nameCol: '', nameFallbacks: [], keyCol: '',
    cands: [], filters: {}, autoDetected: false,
    eligible: [], pool: [], winners: [], winnerIds: new Set(),
    busy: false,
  };

  const nameRes = [
    /^(full[\s_.-]?name|driver[\s_.-]?name|captain[\s_.-]?name)$/i,
    /اسم الكابتن|اسم السائق|الاسم الكامل|اسم الكابتن|الكابتن|السائق/,
    /^(name|الاسم|اسم)$/i,
    /name|اسم|captain|driver/i,
  ];
  const relRe = /مرخص|محرر|رخص|تحرير|licen|releas|free|clear/;
  const negRe = /غير|^لا$|^لا\s|بدون|ليس|\bnot\b|\bno\b|^non|^un|false|^0$|inactive|expired|منتهي|موقوف|محظور|ملغ|مجمد|blocked/;
  const posRe = /مرخص|محرر|رخص|تحرير|licensed|released|cleared|free|true|yes|^1$|نعم|active|فعال|مفعل|valid|ساري|تم/;

  const cell = (r, c) => String(r.data[c] ?? '').trim();
  const isPositive = v => { const s = norm(v); return !!s && !negRe.test(s) && posRe.test(s); };

  function detectNameCol() {
    let best = null;
    for (const c of state.cols) {
      if (c === SRC_COL) continue;
      const idx = nameRes.findIndex(re => re.test(c));
      if (idx < 0) continue;
      const filled = state.rows.reduce((n, r) => n + (cell(r, c) ? 1 : 0), 0);
      if (!best || idx < best.idx || (idx === best.idx && filled > best.filled)) best = { c, idx, filled };
    }
    if (best) return best.c;
    // fallback: first column that looks like text
    for (const c of state.cols) {
      if (c === SRC_COL) continue;
      const vals = state.rows.slice(0, 50).map(r => cell(r, c)).filter(Boolean);
      if (vals.length && vals.filter(v => /[^\d\s+\-().,]/.test(v)).length > vals.length * .7) return c;
    }
    return state.cols[0] || '';
  }

  function detectKeyCol() {
    const j = state.joins.find(x => x.a);
    if (j) return j.a;
    let best = '';
    for (const c of state.cols) {
      if (c === SRC_COL || c === state.nameCol) continue;
      if (/phone|mobile|tel|هاتف|جوال|موبايل|رقم|هوي|معرف|\bid\b|uid/i.test(c)) { best = c; break; }
    }
    return best;
  }

  function buildCandidates(autoSelect) {
    const cands = [];
    for (const c of state.cols) {
      if (c === state.nameCol || c === state.keyCol) continue;
      const counts = new Map();
      for (const r of state.rows) { const v = cell(r, c); counts.set(v, (counts.get(v) || 0) + 1); }
      const headRel = relRe.test(norm(c));
      const valRel = [...counts.keys()].some(v => relRe.test(norm(v)));
      const relevant = headRel || valRel;
      const real = [...counts.keys()].filter(v => v !== '');
      if (counts.size > 25) {
        if (headRel && real.length) cands.push({ col: c, kind: 'nonempty', relevant: true, values: [{ v: '__nonempty__', label: 'لديه قيمة (غير فارغ)', count: state.rows.length - (counts.get('') || 0) }] });
        continue;
      }
      if (counts.size < 2 && !relevant) continue;
      const values = [...counts.entries()].sort((a, b) => b[1] - a[1]).map(([v, count]) => ({ v, label: v === '' ? EMPTY_LABEL : v, count }));
      cands.push({ col: c, kind: 'values', relevant, values });
    }
    cands.sort((a, b) => (b.relevant - a.relevant));
    state.cands = cands;
    if (autoSelect) {
      state.filters = {};
      let any = false;
      for (const cd of cands) {
        if (!cd.relevant) continue;
        const sel = new Set(cd.values.filter(x => x.v === '__nonempty__' || isPositive(x.v)).map(x => x.v));
        if (sel.size) { state.filters[cd.col] = sel; any = true; }
      }
      state.autoDetected = any;
    }
  }

  function computeEligible() {
    const active = state.cands.filter(cd => state.filters[cd.col] && state.filters[cd.col].size);
    const nameOf = r => {
      let v = state.nameCol ? cell(r, state.nameCol) : '';
      if (!v) for (const c of state.nameFallbacks) { v = cell(r, c); if (v) break; }
      if (!v && state.keyCol) v = cell(r, state.keyCol);
      return v || `كابتن #${r.id + 1}`;
    };
    const seen = new Set();
    const list = [];
    for (const r of state.rows) {
      const ok = active.every(cd => {
        const sel = state.filters[cd.col], v = cell(r, cd.col);
        return cd.kind === 'nonempty' ? v !== '' : sel.has(v);
      });
      if (!ok) continue;
      if (state.keyCol) {
        const k = normKey(cell(r, state.keyCol));
        if (k) { if (seen.has(k)) continue; seen.add(k); }
      }
      list.push({ id: r.id, label: nameOf(r), row: r });
    }
    state.eligible = list;
    refreshPool();
  }

  function refreshPool() {
    state.pool = state.eligible.filter(e => !state.winnerIds.has(e.id));
  }

  function recompute({ autoSelect = false } = {}) {
    if (autoSelect) {
      const m = mergeAll(state.datasets);
      state.rows = m.rows; state.cols = m.cols; state.joins = m.joins;
      state.nameCol = detectNameCol();
      state.nameFallbacks = state.cols.filter(c => c !== state.nameCol && c !== SRC_COL && nameRes.slice(0, 3).some(re => re.test(c)));
      state.keyCol = detectKeyCol();
      state.winners = []; state.winnerIds = new Set();
    }
    buildCandidates(autoSelect);
    computeEligible();
    renderAll();
    wheel.resample();
  }

  /* ------------------------------------------------------------------ *
   *  Wheel (canvas)
   * ------------------------------------------------------------------ */
  const wheel = (() => {
    const canvas = $('#wheel');
    const ctx = canvas.getContext('2d');
    const W = { size: 0, dpr: 1, theta: 0, segs: [], anim: null, last: -1, kick: 0, hi: -1, spinning: false, t: 0 };
    const FONT = '"Cairo","Segoe UI",system-ui,sans-serif';
    const Y = '#f9b816', K = '#121213', WH = '#ffffff';
    const PLACEHOLDER = ['TaxiF', 'عمرة 2027', 'TaxiF', '🕋', 'TaxiF', 'عمرة 2027', 'TaxiF', '🕋', 'TaxiF', 'عمرة 2027', 'TaxiF', '🕋'];

    const segColor = (i, n) => (n % 2 === 1 && i === n - 1) ? WH : (i % 2 === 0 ? Y : K);
    const textColor = bg => bg === K ? Y : K;

    function resize() {
      const r = canvas.parentElement.getBoundingClientRect();
      W.size = Math.max(200, r.width);
      W.dpr = Math.min(window.devicePixelRatio || 1, 2.5);
      canvas.width = Math.round(W.size * W.dpr);
      canvas.height = Math.round(W.size * W.dpr);
    }

    function resample() {
      const n = Math.min(state.pool.length, window.innerWidth < 520 ? 14 : 20);
      W.hi = -1; W.last = -1;
      if (!n) { W.segs = PLACEHOLDER.map((label, i) => ({ id: -1 - i, label })); return; }
      W.segs = sample(state.pool, n).map(e => ({ id: e.id, label: e.label }));
    }

    function drawLabel(label, seg, Rs, hubR, color) {
      const rtl = /[؀-ۿ]/.test(label);
      ctx.direction = rtl ? 'rtl' : 'ltr';
      ctx.textAlign = 'right';
      ctx.textBaseline = 'middle';
      const maxW = Rs * .93 - hubR * 1.35;
      const thick = 2 * Math.sin(seg / 2) * (Rs * .55);
      let fs = Math.max(10, Math.min(Rs * .085, thick * .62, 26));
      ctx.font = `800 ${fs}px ${FONT}`;
      let txt = label;
      while (ctx.measureText(txt).width > maxW && fs > 11) { fs -= 1; ctx.font = `800 ${fs}px ${FONT}`; }
      if (ctx.measureText(txt).width > maxW) {
        while (txt.length > 2 && ctx.measureText(txt + '…').width > maxW) txt = txt.slice(0, -1);
        txt = txt.trimEnd() + '…';
      }
      ctx.fillStyle = color;
      ctx.fillText(txt, Rs * .93, 0);
    }

    function render(now) {
      const S = W.size, C = S / 2, R = C * .975, rim = S * .052, Rs = R - rim, hubR = Rs * .19;
      const n = W.segs.length, seg = TAU / n;
      ctx.setTransform(W.dpr, 0, 0, W.dpr, 0, 0);
      ctx.clearRect(0, 0, S, S);

      // outer disc
      ctx.save();
      ctx.shadowColor = 'rgba(249,184,22,.45)'; ctx.shadowBlur = S * .06;
      ctx.beginPath(); ctx.arc(C, C, R, 0, TAU);
      const g = ctx.createLinearGradient(0, 0, S, S);
      g.addColorStop(0, '#2a2a2c'); g.addColorStop(.5, '#0b0b0c'); g.addColorStop(1, '#2a2a2c');
      ctx.fillStyle = g; ctx.fill();
      ctx.restore();
      ctx.beginPath(); ctx.arc(C, C, R - 1, 0, TAU);
      ctx.strokeStyle = 'rgba(249,184,22,.6)'; ctx.lineWidth = 2; ctx.stroke();

      // segments
      ctx.save();
      ctx.translate(C, C);
      ctx.rotate(W.theta);
      for (let i = 0; i < n; i++) {
        const a0 = -Math.PI / 2 + i * seg, a1 = a0 + seg;
        const bg = segColor(i, n);
        ctx.beginPath(); ctx.moveTo(0, 0); ctx.arc(0, 0, Rs, a0, a1); ctx.closePath();
        ctx.fillStyle = bg; ctx.fill();
        if (i === W.hi) {
          const pulse = .35 + .35 * Math.sin(now / 130);
          ctx.fillStyle = `rgba(255,255,255,${pulse})`; ctx.fill();
        }
        ctx.strokeStyle = 'rgba(0,0,0,.35)'; ctx.lineWidth = 1.5; ctx.stroke();
      }
      for (let i = 0; i < n; i++) {
        ctx.save();
        ctx.rotate(-Math.PI / 2 + (i + .5) * seg);
        drawLabel(W.segs[i].label, seg, Rs, hubR, textColor(segColor(i, n)));
        ctx.restore();
      }
      ctx.restore();

      // winning segment glow ring on the rim
      ctx.beginPath(); ctx.arc(C, C, Rs, 0, TAU);
      ctx.strokeStyle = '#0b0b0c'; ctx.lineWidth = 3; ctx.stroke();

      // bulbs
      const bulbs = 28, br = rim * .2, bR = R - rim / 2;
      const phase = Math.floor(now / (W.spinning ? 80 : 450));
      for (let i = 0; i < bulbs; i++) {
        const a = i / bulbs * TAU, on = (i + phase) % 2 === 0;
        ctx.save();
        ctx.beginPath(); ctx.arc(C + Math.cos(a) * bR, C + Math.sin(a) * bR, br, 0, TAU);
        if (on) { ctx.shadowColor = '#ffd35c'; ctx.shadowBlur = br * 3; }
        ctx.fillStyle = on ? '#ffe08a' : '#6f5412';
        ctx.fill();
        ctx.restore();
      }

      // hub backdrop
      ctx.beginPath(); ctx.arc(C, C, hubR * 1.22, 0, TAU);
      ctx.fillStyle = '#0b0b0c'; ctx.fill();

      // pointer
      const px = C, py = C - R + rim * .25;
      ctx.save();
      ctx.translate(px, py);
      ctx.rotate(-W.kick * .38);
      ctx.shadowColor = 'rgba(0,0,0,.6)'; ctx.shadowBlur = S * .02; ctx.shadowOffsetY = S * .008;
      ctx.beginPath();
      ctx.moveTo(-S * .042, -S * .014); ctx.lineTo(S * .042, -S * .014); ctx.lineTo(0, S * .098); ctx.closePath();
      ctx.fillStyle = Y; ctx.fill();
      ctx.shadowColor = 'transparent';
      ctx.lineWidth = S * .008; ctx.strokeStyle = '#0b0b0c'; ctx.lineJoin = 'round'; ctx.stroke();
      ctx.beginPath(); ctx.arc(0, 0, S * .012, 0, TAU); ctx.fillStyle = '#0b0b0c'; ctx.fill();
      ctx.restore();
    }

    function frame(now) {
      const dt = Math.min(.05, (now - (W.t || now)) / 1000); W.t = now;
      if (W.anim) {
        const a = W.anim, p = Math.min(1, (now - a.t0) / a.dur);
        W.theta = a.from + a.delta * (1 - Math.pow(1 - p, 4));
        if (p >= 1) { const cb = a.done; W.anim = null; W.spinning = false; cb && cb(); }
      } else {
        W.theta += dt * .1;
      }
      const n = W.segs.length, seg = TAU / n;
      const idx = Math.floor(mod(-W.theta, TAU) / seg);
      if (idx !== W.last) {
        if (W.last !== -1 && W.spinning) { W.kick = 1; sound.tick(); }
        W.last = idx;
      }
      W.kick *= Math.exp(-dt * 10);
      render(now);
      requestAnimationFrame(frame);
    }

    new ResizeObserver(resize).observe(canvas.parentElement);
    resize(); resample();
    requestAnimationFrame(frame);
    if (document.fonts && document.fonts.load) document.fonts.load('800 20px Cairo').catch(() => {});

    return {
      resample,
      get spinning() { return W.spinning; },
      /** Spins and lands on `win`. */
      spin(win, onLanded) {
        const n = W.segs.length, seg = TAU / n;
        let w = W.segs.findIndex(s => s.id === win.id);
        if (w < 0) { w = randInt(n); W.segs[w] = { id: win.id, label: win.label }; }
        const off = (Math.random() - .5) * seg * .6;
        W.theta = mod(W.theta, TAU);
        let target = -((w + .5) * seg + off);
        const minEnd = W.theta + TAU * (6 + randInt(3));
        target += Math.ceil((minEnd - target) / TAU) * TAU;
        W.hi = -1; W.spinning = true;
        W.anim = { from: W.theta, delta: target - W.theta, t0: performance.now(), dur: 7200, done: () => { W.hi = w; onLanded(); } };
      },
      clearHighlight() { W.hi = -1; },
    };
  })();

  /* ------------------------------------------------------------------ *
   *  Rendering the side panels
   * ------------------------------------------------------------------ */
  const el = id => document.getElementById(id);

  function maskId(v) {
    const s = String(v || '').trim();
    if (!s) return '';
    return s.length <= 4 ? '••' + s.slice(-2) : '••••' + s.slice(-4);
  }

  function renderFiles() {
    const ul = el('fileList'); ul.textContent = '';
    state.datasets.forEach((d, i) => {
      const li = document.createElement('li');
      const tag = document.createElement('span'); tag.className = 'ftag' + (d.ext === 'json' ? '' : ' xl'); tag.textContent = d.ext === 'json' ? 'JSON' : 'XLSX';
      const meta = document.createElement('div'); meta.className = 'fmeta';
      const b = document.createElement('b'); b.textContent = d.name;
      const s = document.createElement('small'); s.textContent = `${fmt(d.rows.length)} سجل · ${d.columns.length} عمود`;
      meta.append(b, s);
      const x = document.createElement('button'); x.className = 'x'; x.type = 'button'; x.textContent = '×'; x.title = 'إزالة'; x.dataset.i = i;
      li.append(tag, meta, x); ul.append(li);
    });
    el('filesEmpty').hidden = state.datasets.length > 0;
    const info = el('joinInfo');
    const j = state.joins.find(x => x.a);
    if (j) {
      info.hidden = false;
      info.textContent = `🔗 تم ربط الملفين تلقائياً عبر «${j.a}» ↔ «${j.b}» (${fmt(j.matched)} تطابق).`;
    } else if (state.joins.some(x => x.none)) {
      info.hidden = false;
      info.textContent = '⚠️ لم أجد عموداً مشتركاً لربط الملفين، فتم دمج السجلات بعضها فوق بعض.';
    } else info.hidden = true;
  }

  function renderRules() {
    const card = el('rulesCard');
    card.hidden = !state.datasets.length;
    if (!state.datasets.length) return;
    const host = el('rules'); host.textContent = '';
    const mkRule = (cd, idx) => {
      const div = document.createElement('div'); div.className = 'rule';
      const t = document.createElement('div'); t.className = 'rule-t' + (cd.relevant ? ' rel' : ''); t.textContent = cd.col;
      const chips = document.createElement('div'); chips.className = 'chips';
      for (const v of cd.values) {
        const c = document.createElement('button'); c.type = 'button'; c.className = 'chip';
        c.dataset.col = idx; c.dataset.val = v.v;
        c.setAttribute('aria-pressed', String(!!(state.filters[cd.col] && state.filters[cd.col].has(v.v))));
        c.append(document.createTextNode(v.label));
        const sm = document.createElement('small'); sm.textContent = fmt(v.count); c.append(sm);
        chips.append(c);
      }
      div.append(t, chips);
      return div;
    };
    const main = state.cands.map((cd, i) => [cd, i]).filter(([cd]) => cd.relevant);
    const rest = state.cands.map((cd, i) => [cd, i]).filter(([cd]) => !cd.relevant);
    main.forEach(([cd, i]) => host.append(mkRule(cd, i)));
    if (rest.length) {
      const d = document.createElement('details'); d.className = 'more';
      const s = document.createElement('summary'); s.textContent = `أعمدة أخرى (${rest.length})`;
      d.append(s); rest.forEach(([cd, i]) => d.append(mkRule(cd, i)));
      d.open = !main.length;
      host.append(d);
    }
    if (!state.cands.length) { const p = document.createElement('p'); p.className = 'muted small'; p.textContent = 'لا توجد أعمدة قابلة للتصفية — سيدخل الجميع في السحب.'; host.append(p); }
    const active = Object.values(state.filters).filter(s => s.size).length;
    el('rulesCount').textContent = active ? `${active} شرط فعّال` : 'بدون شروط';
    el('rulesNotice').hidden = state.autoDetected || active > 0;

    const fillSel = (sel, cur, withNone) => {
      sel.textContent = '';
      if (withNone) { const o = document.createElement('option'); o.value = ''; o.textContent = '— بدون —'; sel.append(o); }
      for (const c of state.cols.filter(c => c !== SRC_COL)) { const o = document.createElement('option'); o.value = c; o.textContent = c; sel.append(o); }
      sel.value = cur;
    };
    fillSel(el('nameSel'), state.nameCol, false);
    fillSel(el('keySel'), state.keyCol, true);
  }

  function renderWinners() {
    const ol = el('winnerList'); ol.textContent = '';
    for (const w of state.winners) {
      const li = document.createElement('li');
      const n = document.createElement('span'); n.className = 'n'; n.textContent = w.n;
      const nm = document.createElement('span'); nm.className = 'w'; nm.textContent = w.label;
      const id = document.createElement('span'); id.className = 'id'; id.textContent = maskId(w.row.data[state.keyCol]);
      li.append(n, nm, id); ol.append(li);
    }
    el('winnersEmpty').hidden = state.winners.length > 0;
    el('exportBtn').disabled = !state.winners.length;
    el('resetWinners').hidden = !state.winners.length;
  }

  function renderStatus() {
    el('statEligible').textContent = fmt(state.pool.length);
    el('statTotal').textContent = fmt(state.rows.length);
    el('statWinners').textContent = fmt(state.winners.length);
    el('empty').hidden = state.datasets.length > 0;
    const btn = el('spinBtn'), hub = el('hub'), lbl = btn.querySelector('.spin-label');
    const ready = state.pool.length > 0 && !state.busy;
    btn.disabled = !ready; hub.disabled = !ready;
    hub.hidden = !state.datasets.length;
    hub.classList.toggle('ready', ready);
    btn.classList.toggle('busy', state.busy);
    lbl.textContent = state.busy ? 'جارٍ السحب…'
      : !state.datasets.length ? 'حمّل الملفات أولاً'
      : !state.pool.length ? (state.eligible.length ? 'اكتمل السحب — لا يوجد مؤهلون آخرون' : 'لا يوجد كباتن مؤهلون')
      : (state.winners.length ? 'اسحب فائزاً آخر 🎯' : 'ابدأ السحب 🎯');
  }

  function renderAll() { renderFiles(); renderRules(); renderWinners(); renderStatus(); }

  /* ------------------------------------------------------------------ *
   *  Loading files
   * ------------------------------------------------------------------ */
  async function addFiles(fileList) {
    const files = [...fileList];
    if (!files.length) return;
    let added = 0;
    for (const f of files) {
      try {
        const ds = await readFile(f);
        const i = state.datasets.findIndex(d => d.name === ds.name);
        if (i >= 0) state.datasets[i] = ds; else state.datasets.push(ds);
        added++;
      } catch (e) { toast(e.message || 'تعذّر قراءة الملف', true); }
    }
    if (!added) return;
    recompute({ autoSelect: true });
    const n = state.pool.length;
    toast(`تم تحميل ${added} ملف · ${fmt(n)} كابتن مؤهل`);
  }

  function loadDemo() {
    const first = ['محمد', 'أحمد', 'خالد', 'عمر', 'يوسف', 'ياسر', 'سامر', 'طارق', 'ماهر', 'بسام', 'فادي', 'رامي', 'وسيم', 'نبيل', 'هاني', 'زياد', 'علي', 'حسن', 'ماجد', 'سعيد'];
    const last = ['الحلبي', 'الخطيب', 'النجار', 'الشامي', 'العلي', 'الأحمد', 'حداد', 'قاسم', 'ديب', 'المصري', 'عيسى', 'سلامة', 'درويش', 'الحسن', 'صالح'];
    const rows = [];
    for (let i = 0; i < 90; i++) {
      rows.push({
        'اسم الكابتن': `${first[randInt(first.length)]} ${last[randInt(last.length)]}`,
        'رقم الهاتف': '09' + String(10000000 + randInt(89999999)),
        'حالة الترخيص': randInt(100) < 75 ? 'مرخص' : 'غير مرخص',
        'حالة التحرير': randInt(100) < 70 ? 'محرر' : 'غير محرر',
      });
    }
    state.datasets = [{ name: 'بيانات-تجريبية.json', ext: 'json', rows, columns: columnsOf(rows) }];
    recompute({ autoSelect: true });
    toast('تم تحميل بيانات تجريبية — اسحب ملفاتك الحقيقية لاستبدالها');
  }

  /* ------------------------------------------------------------------ *
   *  Spin + celebration
   * ------------------------------------------------------------------ */
  function startSpin() {
    if (state.busy || !state.pool.length) return;
    sound.unlock();
    const win = state.pool[randInt(state.pool.length)];
    state.busy = true;
    renderStatus();
    wheel.spin(win, () => {
      state.winners.push({ n: state.winners.length + 1, id: win.id, label: win.label, row: win.row });
      state.winnerIds.add(win.id);
      refreshPool();
      sound.win();
      renderWinners(); renderStatus();
      setTimeout(() => showWinner(state.winners[state.winners.length - 1]), 900);
    });
  }

  function celebrate() {
    if (typeof confetti !== 'function') return;
    const colors = ['#f9b816', '#ffd35c', '#ffffff', '#fff1b8', '#111111'];
    const z = 3000;
    confetti({ particleCount: 180, spread: 110, startVelocity: 55, origin: { x: .5, y: .62 }, colors, zIndex: z, scalar: 1.1 });
    confetti({ particleCount: 60, spread: 360, startVelocity: 25, gravity: .6, ticks: 260, origin: { x: .5, y: .4 }, shapes: ['star'], colors, zIndex: z, scalar: 1.4 });
    try {
      const kaaba = confetti.shapeFromText({ text: '🕋', scalar: 3 });
      confetti({ particleCount: 14, spread: 100, startVelocity: 38, origin: { x: .5, y: .7 }, shapes: [kaaba], scalar: 3, zIndex: z, flat: true });
    } catch { /* emoji shapes are optional */ }
    const end = Date.now() + 5000;
    (function side() {
      confetti({ particleCount: 5, angle: 60, spread: 70, origin: { x: 0, y: .75 }, colors, zIndex: z });
      confetti({ particleCount: 5, angle: 120, spread: 70, origin: { x: 1, y: .75 }, colors, zIndex: z });
      if (Date.now() < end && !el('winModal').hidden) requestAnimationFrame(side);
    })();
    setTimeout(() => { if (!el('winModal').hidden) confetti({ particleCount: 90, spread: 140, startVelocity: 40, origin: { x: .5, y: .3 }, colors, zIndex: z }); }, 1800);
  }

  function showWinner(w) {
    el('winName').textContent = w.label;
    el('winNum').textContent = `الفائز رقم ${w.n}`;
    el('winId').textContent = maskId(w.row.data[state.keyCol]);
    el('winModal').hidden = false;
    celebrate();
    el('nextBtn').focus({ preventScroll: true });
  }

  function closeWinner(thenSpin) {
    el('winModal').hidden = true;
    wheel.clearHighlight();
    wheel.resample();
    state.busy = false;
    renderStatus();
    if (thenSpin) setTimeout(startSpin, 250);
  }

  /* ------------------------------------------------------------------ *
   *  Export
   * ------------------------------------------------------------------ */
  function exportWinners() {
    if (!window.XLSX || !state.winners.length) return;
    const rows = state.winners.map(w => ({ '#': w.n, 'الاسم': w.label, ...Object.fromEntries(Object.entries(w.row.data).filter(([k]) => k !== SRC_COL)) }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Views: [{ RTL: true }] };
    XLSX.utils.book_append_sheet(wb, ws, 'الفائزون');
    XLSX.writeFile(wb, 'taxif-omra-2027-winners.xlsx');
  }

  /* ------------------------------------------------------------------ *
   *  Events
   * ------------------------------------------------------------------ */
  const input = el('fileInput');
  const pick = () => { input.value = ''; input.click(); };
  input.addEventListener('change', () => addFiles(input.files));
  el('pickBtn').addEventListener('click', e => { e.stopPropagation(); pick(); });
  el('addBtn').addEventListener('click', pick);
  el('demoBtn').addEventListener('click', e => { e.stopPropagation(); loadDemo(); });
  el('dropZone').addEventListener('click', pick);
  el('dropZone').addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); } });

  // drag & drop anywhere on the page
  let dragDepth = 0;
  const hasFiles = e => e.dataTransfer && [...(e.dataTransfer.types || [])].includes('Files');
  window.addEventListener('dragenter', e => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; document.body.classList.add('dragging'); });
  window.addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('dragleave', e => { if (!hasFiles(e)) return; dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) document.body.classList.remove('dragging'); });
  window.addEventListener('drop', e => {
    if (!hasFiles(e)) return;
    e.preventDefault(); dragDepth = 0; document.body.classList.remove('dragging');
    addFiles(e.dataTransfer.files);
  });

  el('fileList').addEventListener('click', e => {
    const b = e.target.closest('.x'); if (!b) return;
    state.datasets.splice(Number(b.dataset.i), 1);
    if (state.datasets.length) recompute({ autoSelect: true });
    else {
      Object.assign(state, { rows: [], cols: [], joins: [], cands: [], filters: {}, eligible: [], pool: [], winners: [], winnerIds: new Set(), nameCol: '', keyCol: '' });
      renderAll(); wheel.resample();
    }
  });

  el('rules').addEventListener('click', e => {
    const chip = e.target.closest('.chip'); if (!chip) return;
    const cd = state.cands[Number(chip.dataset.col)];
    const set = state.filters[cd.col] || (state.filters[cd.col] = new Set());
    const v = chip.dataset.val;
    set.has(v) ? set.delete(v) : set.add(v);
    chip.setAttribute('aria-pressed', String(set.has(v)));
    computeEligible();
    el('rulesCount').textContent = Object.values(state.filters).filter(s => s.size).length + ' شرط فعّال';
    el('rulesNotice').hidden = true;
    renderWinners(); renderStatus(); wheel.resample();
  });
  el('nameSel').addEventListener('change', e => { state.nameCol = e.target.value; recompute(); });
  el('keySel').addEventListener('change', e => { state.keyCol = e.target.value; recompute(); });

  el('spinBtn').addEventListener('click', startSpin);
  el('hub').addEventListener('click', startSpin);
  el('nextBtn').addEventListener('click', () => closeWinner(true));
  el('closeBtn').addEventListener('click', () => closeWinner(false));
  el('exportBtn').addEventListener('click', exportWinners);
  el('resetWinners').addEventListener('click', () => {
    state.winners = []; state.winnerIds = new Set(); refreshPool(); renderWinners(); renderStatus(); wheel.resample();
  });

  el('soundBtn').addEventListener('click', e => {
    sound.muted = !sound.muted;
    e.currentTarget.setAttribute('aria-pressed', String(!sound.muted));
  });

  window.addEventListener('keydown', e => {
    const modalOpen = !el('winModal').hidden;
    if (e.key === 'Escape' && modalOpen) closeWinner(false);
    if (e.code === 'Space' && !/^(INPUT|SELECT|TEXTAREA|BUTTON|SUMMARY)$/.test(document.activeElement.tagName)) {
      e.preventDefault();
      modalOpen ? closeWinner(true) : startSpin();
    }
  });

  let rsz;
  window.addEventListener('resize', () => { clearTimeout(rsz); rsz = setTimeout(() => { if (!wheel.spinning) wheel.resample(); }, 250); });

  renderAll();

  // test hook (harmless in production)
  window.__draw = { state, addFiles, loadDemo, startSpin };
})();
