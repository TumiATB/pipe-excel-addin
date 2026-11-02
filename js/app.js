/* Pipe task pane: engagement setup, full audit and standalone processes, chat. */
(function () {
  const D = PipeData, WP = PipeWP;
  const $ = s => document.querySelector(s);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const R = n => 'R ' + Math.round(n).toLocaleString('en-US');
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const now = () => new Date().toTimeString().slice(0, 5);
  const num = (v, d) => typeof v === 'number' && isFinite(v) ? v : d;
  let XL, inExcel = false;

  const UNITS = {
    wf1: {name:'GL analysis', procs:['mat','je','risk'], next:'wf2', n:1},
    wf2: {name:'Sampling', procs:['samp'], next:'wf3', n:2},
    wf3: {name:'Detail testing', procs:['test','sum'], next:null, n:3}
  };
  const PROCS = {
    mat: {name:'Materiality', ref:'A1', sheets:['A1 Materiality'], needs:['tb','basis'], makes:['mat'], build:WP.materiality},
    je: {name:'Journal testing', ref:'A3', sheets:['A3 Journals'], needs:['gl','users'], makes:['je'], build:WP.journals},
    risk: {name:'Risk assessment', ref:'A2', sheets:['A2 Risk assessment'], needs:['mat'], makes:['risks'], build:WP.risk},
    samp: {name:'Sampling', ref:'B1', sheets:['B1 Sampling'], needs:['tb','mat','risks'], makes:['samples'], build:WP.sampling},
    test: {name:'Detail testing', ref:'D3.1–J1.1', sheets:D.AREAS.map(a => a.sheet), needs:['samples','mat','docs'], makes:['results'], build:WP.detailAll},
    sum: {name:'Misstatements', ref:'S1', sheets:['S1 Misstatements'], needs:['results','mat'], makes:[], build:WP.summary}
  };
  const NEED = {tb:'Trial balance', gl:'General ledger', basis:'Materiality basis', users:'Senior users', mat:'Materiality', risks:'Risk levels', samples:'Samples', docs:'Supporting documents', results:'Detail testing results'};
  const DEFAULT_SIZES = Object.fromEntries(D.AREAS.map(a => [a.id, a.size]));
  const DEFAULT_RISK = {REV:'Significant', INV:'High', PAY:'Medium', PPE:'Medium', JE:'Significant'};

  /* ---------- State ---------- */
  const fresh = () => ({setup:null, units:{wf1:{status:'ready'}, wf2:{status:'locked'}, wf3:{status:'locked'}}, sizes:{...DEFAULT_SIZES}, wp:{}, docs:false, tab:'audit', open:null,
    chat:[{who:'ai', text:'Hi. I write audit working papers straight into this workbook. Set up the engagement, then run the full audit or any single process.'}]});
  let ST = fresh();
  const U = uid => ST.units[uid] || (ST.units[uid] = {status:'ready'});
  const procsOf = uid => UNITS[uid] ? UNITS[uid].procs : [uid.slice(2)];
  const nameOf = uid => UNITS[uid] ? UNITS[uid].name : PROCS[uid.slice(2)].name;
  function needsOf(uid) {
    const made = new Set(), out = [];
    procsOf(uid).forEach(p => { PROCS[p].needs.forEach(n => { if (!made.has(n) && !out.includes(n)) out.push(n); }); PROCS[p].makes.forEach(m => made.add(m)); });
    return out;
  }
  function save() {
    const data = JSON.stringify({...ST, chat:ST.chat.slice(-40)});
    try {
      if (inExcel) { Office.context.document.settings.set('pipe', data); Office.context.document.settings.saveAsync(() => {}); }
      else localStorage.setItem('pipe-addin', data);
    } catch (e) {}
  }
  function load() {
    try {
      const raw = inExcel ? Office.context.document.settings.get('pipe') : localStorage.getItem('pipe-addin');
      if (raw) ST = Object.assign(fresh(), JSON.parse(raw));
    } catch (e) {}
    Object.values(ST.units).forEach(u => { if (['checking','running'].includes(u.status)) u.status = u.status === 'running' ? 'review' : 'ready'; });
  }

  /* ---------- Context detection ---------- */
  async function readNum(sheet, a, d) { try { return num((await XL.read(sheet, a))[0][0], d); } catch (e) { return d; } }
  async function detect(key) {
    const s = ST.setup;
    switch (key) {
      case 'tb': return await XL.exists(s.tbSheet) ? `Sheet "${s.tbSheet}"` : null;
      case 'gl': return await XL.exists(s.glSheet) ? `Sheet "${s.glSheet}"` : null;
      case 'basis': return `${WP.BASIS[s.entity]} · setup`;
      case 'users': return `${s.users} · setup`;
      case 'mat': return await XL.exists('A1 Materiality') ? `A1 · ${R(await readNum('A1 Materiality', 'J17', 740000))}` : null;
      case 'risks': return await XL.exists('A2 Risk assessment') ? 'A2 risk register' : null;
      case 'samples': return await XL.exists('B1 Sampling') ? 'B1 selected items' : null;
      case 'docs': return ST.docs === true ? '224 documents indexed' : null;
      case 'results': { for (const a of D.AREAS) if (!await XL.exists(a.sheet)) return null; return 'D3.1 to J1.1'; }
    }
    return null;
  }
  async function ctxFor(u) {
    const s = ST.setup, A = u.asked || {}, ex = n => XL.exists(n);
    const C = {tbName:s.tbSheet, glName:s.glSheet, basis:s.entity, sizes:{...ST.sizes}, risks:A.risks || DEFAULT_RISK, manual:A.manual, det:{}};
    C.tb = await ex(s.tbSheet); C.gl = await ex(s.glSheet);
    C.a1 = await ex('A1 Materiality'); C.a2 = await ex('A2 Risk assessment'); C.a3 = await ex('A3 Journals'); C.b1 = await ex('B1 Sampling');
    for (const a of D.AREAS) C.det[a.id] = await ex(a.sheet);
    C.om = C.a1 ? await readNum('A1 Materiality', 'J17', 740000) : num(A.om, 740000);
    C.pm = C.a1 ? await readNum('A1 Materiality', 'J19', 555000) : num(A.pm, C.om * 0.75);
    C.ct = C.a1 ? await readNum('A1 Materiality', 'J21', 37000) : C.om * 0.05;
    if (C.b1) {
      try { const v = await XL.read('B1 Sampling', 'G29:G33'); D.AREAS.forEach((a, i) => { C.sizes[a.id] = num(v[i][0], C.sizes[a.id]); }); } catch (e) {}
    }
    C.samples = D.samples(C.sizes);
    C.bench = {revenue:186.4e6, pbt:14.87e6, equity:63.65e6, assets:130.32e6, liabilities:51.8e6};
    return C;
  }
  const dummyCtx = () => ({tbName:'TB', glName:'GL', basis:(ST.setup || {}).entity, sizes:ST.sizes, risks:DEFAULT_RISK, det:{}, tb:true, gl:true, samples:D.samples(ST.sizes), om:740000, pm:555000, ct:37000});

  /* ---------- Writing ---------- */
  async function write(sec) {
    const o = XLUtil.ops(); sec.ops(o);
    await XL.apply(sec.sheet, o.list, {clear:sec.clear, tab:sec.tab, pos:WP.posFor(sec.sheet), select:sec.select, quiet:sec.quiet});
  }
  async function writeAll(secs) { for (const s of secs) await write(s); }
  const refOf = sheet => WP.WPS.find(w => w[1] === sheet)[0];
  async function markIndex(sheets, status, which) {
    if (!await XL.exists('Index')) return;
    const o = XLUtil.ops();
    sheets.forEach(sh => { const ref = refOf(sh); WP.indexRow(ref, status, which)(o); ST.wp[ref] = {...ST.wp[ref], status, [which]:WP.today()}; });
    await XL.apply('Index', o.list, {quiet:true});
  }

  /* ---------- Engine ---------- */
  async function check(uid) {
    const u = U(uid);
    u.status = 'checking'; u.asked = {}; draw();
    u.needs = [];
    for (const k of needsOf(uid)) u.needs.push({k, found:await detect(k)});
    u.sheetsList = (await XL.sheets()).filter(n => !WP.ORDER.slice(0, 11).includes(n));
    u.status = 'ctx'; draw();
  }
  async function confirmCtx(uid) {
    const u = U(uid), A = u.asked = {}, val = id => { const el = document.getElementById(id); return el ? el.value : null; };
    for (const n of u.needs) {
      if (n.found) continue;
      if (n.k === 'tb' || n.k === 'gl') {
        const v = val(`ask-${uid}-${n.k}`);
        if (!v) { toast(`Choose a ${NEED[n.k].toLowerCase()} sheet or create sample data`); return; }
        ST.setup[n.k === 'tb' ? 'tbSheet' : 'glSheet'] = v;
      }
      if (n.k === 'mat') { A.om = +val(`ask-${uid}-om`); A.pm = +val(`ask-${uid}-pm`); }
      if (n.k === 'risks') A.risks = Object.fromEntries(D.AREAS.map(a => [a.id, val(`ask-${uid}-r-${a.id}`)]));
      if (n.k === 'docs') { toast('Request the supporting documents first'); return; }
      if (n.k === 'results') A.manual = Object.fromEntries(D.AREAS.map(a => [a.id, {known:+val(`ask-${uid}-k-${a.id}`), projected:+val(`ask-${uid}-p-${a.id}`), exc:+val(`ask-${uid}-e-${a.id}`)}]));
    }
    const asked = u.needs.filter(n => !n.found).map(n => NEED[n.k].toLowerCase());
    u.status = 'plan'; draw();
    say(asked.length ? `Got it. I'll use the ${asked.join(', ')} you gave me and mark ${asked.length > 1 ? 'them' : 'it'} as auditor-provided (shaded) in the paper.` : `All inputs found in the workbook. Here's the plan for ${nameOf(uid).toLowerCase()}.`);
  }
  async function run(uid) {
    const u = U(uid);
    u.status = 'running'; u.done = []; u.cur = null; u.err = null; draw();
    try {
      for (const p of procsOf(uid)) {
        const C = await ctxFor(u);
        for (const sec of PROCS[p].build(C)) {
          u.cur = sec.step; draw();
          await write(sec);
          u.done.push(sec.step); draw();
          await sleep(inExcel ? 120 : 220);
        }
        await markIndex(PROCS[p].sheets, 'Prepared', 'prep');
      }
      u.cur = null; u.status = 'review'; u.results = await results(uid);
      draw();
      say(`${nameOf(uid)} is written to the workbook. ${u.results.join(' ')} Check the shaded judgement cells, then sign off.`);
    } catch (e) {
      u.status = 'plan'; u.err = e.message || String(e); draw();
    }
  }
  async function results(uid) {
    const out = [];
    for (const p of procsOf(uid)) {
      if (p === 'mat') out.push(`Materiality ${R(await readNum('A1 Materiality', 'J17', 740000))}, PM ${R(await readNum('A1 Materiality', 'J19', 555000))}.`);
      if (p === 'je') out.push('3 journals scored 3+ (JE-24117, JE-24203, JE-23980).');
      if (p === 'risk') out.push('2 significant risks (revenue, journals) and 1 high (inventory).');
      if (p === 'samp') out.push(`${Object.values(ST.sizes).reduce((a, b) => a + b, 0)} items selected across 5 areas.`);
      if (p === 'test') { const s = D.samples(ST.sizes); let e = 0, m = 0; D.AREAS.forEach(a => { const r = D.evalArea(a, s.filter(x => x.area === a.id)); e += r.exc; m += r.missing; }); out.push(`${e} exceptions and ${m} without support.`); }
      if (p === 'sum') { const s = D.samples(ST.sizes); let k = 0, pr = 0; D.AREAS.forEach(a => { const r = D.evalArea(a, s.filter(x => x.area === a.id)); k += r.known; pr += r.projected; }); out.push(`Known ${R(k)}, projected ${R(pr)}.`); }
    }
    return out;
  }
  async function signOff(uid) {
    const u = U(uid), sheets = procsOf(uid).flatMap(p => PROCS[p].sheets);
    for (const sh of sheets) if (await XL.exists(sh)) { const o = XLUtil.ops(); WP.signoff(ST.setup.reviewer)(o); await XL.apply(sh, o.list, {quiet:true}); }
    await markIndex(sheets, 'Reviewed', 'rev');
    u.status = 'done'; u.signed = `${ST.setup.reviewer} · ${new Date().toLocaleDateString('en-GB', {day:'numeric', month:'short'})}`;
    const nx = UNITS[uid] && UNITS[uid].next;
    if (nx && U(nx).status === 'locked') U(nx).status = 'ready';
    draw();
    toast(`${nameOf(uid)} signed off`);
    say(`${nameOf(uid)} signed off. Reviewer initials and date are in each paper's header and on the Index.${nx ? ` ${UNITS[nx].name} is unlocked.` : ''}`);
  }
  async function requestDocs() {
    ST.docs = 0; draw();
    for (let p = 10; p <= 100; p += 15) { await sleep(220); ST.docs = Math.min(p, 99); draw(); }
    ST.docs = true;
    Object.keys(ST.units).forEach(id => { const u = ST.units[id]; if (u.needs) u.needs.forEach(n => { if (n.k === 'docs') n.found = '224 documents indexed'; }); });
    draw(); say('224 documents came back from the client portal and are indexed.');
  }

  /* ---------- Setup ---------- */
  async function saveSetup() {
    const v = id => document.getElementById(id).value.trim();
    const set = {client:v('s-client'), ye:v('s-ye'), framework:v('s-fw'), entity:v('s-entity'), reviewer:v('s-rev').toUpperCase(), users:v('s-users'), tbSheet:v('s-tb') || 'TB', glSheet:v('s-gl') || 'GL'};
    if (!set.client || !set.ye || !set.reviewer) { toast('Client, year end and reviewer are required'); return; }
    ST.setup = set; ST.view = null;
    busy(true);
    try { await writeAll(WP.index(set, ST.wp)); } catch (e) { toast(e.message); }
    busy(false);
    draw(); say(`Engagement saved for ${set.client}. Every working paper header links to the Index sheet.`);
  }
  async function sampleData() {
    busy(true);
    try { await writeAll(WP.sampleData()); toast('Created TB and GL sheets'); } catch (e) { toast(e.message); }
    busy(false);
    Object.values(ST.units).forEach(u => { if (u.status === 'ctx') u.status = 'ready'; });
    draw();
  }

  /* ---------- UI ---------- */
  const PILL = {locked:['','Locked'], ready:['info','Ready'], checking:['ai','Checking'], ctx:['ai','Inputs'], plan:['warn','Plan'], running:['ai','Running'], review:['warn','Review'], done:['ok','Signed off']};
  const pill = st => `<span class="pill ${PILL[st][0]}">${PILL[st][1]}</span>`;
  function needRow(uid, n) {
    if (n.found) return `<li class="ok"><i>✓</i><div><b>${NEED[n.k]}</b><span>${esc(n.found)}</span></div></li>`;
    const u = U(uid), id = `ask-${uid}`;
    let w = '';
    if (n.k === 'tb' || n.k === 'gl') w = `<div class="row"><select id="${id}-${n.k}"><option value="">Choose sheet…</option>${(u.sheetsList || []).map(s => `<option>${esc(s)}</option>`).join('')}</select><button class="btn sm" data-act="sample">Create sample data</button></div>`;
    if (n.k === 'mat') w = `<div class="row"><label>Materiality<input id="${id}-om" type="number" value="740000" step="10000"></label><label>Performance<input id="${id}-pm" type="number" value="555000" step="5000"></label></div>`;
    if (n.k === 'risks') w = `<div class="grid5">${D.AREAS.map(a => `<label>${a.name}<select id="${id}-r-${a.id}">${['Significant','High','Medium','Low'].map(l => `<option ${l === DEFAULT_RISK[a.id] ? 'selected' : ''}>${l}</option>`).join('')}</select></label>`).join('')}</div>`;
    if (n.k === 'samples') w = `<span class="hint">None in this workbook. Pipe will draw its default selection.</span>`;
    if (n.k === 'docs') w = ST.docs === false || ST.docs == null ? `<button class="btn sm" data-act="docs">Request from client portal</button>` : `<div class="bar"><i style="width:${ST.docs === true ? 100 : ST.docs}%"></i></div>`;
    if (n.k === 'results') { const s = D.samples(ST.sizes); w = `<table class="mini"><tr><th></th><th>Exc.</th><th>Known</th><th>Projected</th></tr>${D.AREAS.map(a => { const r = D.evalArea(a, s.filter(x => x.area === a.id)); return `<tr><td>${a.ref}</td><td><input id="${id}-e-${a.id}" type="number" value="${r.exc + r.missing}"></td><td><input id="${id}-k-${a.id}" type="number" value="${Math.round(r.known)}"></td><td><input id="${id}-p-${a.id}" type="number" value="${Math.round(r.projected)}"></td></tr>`; }).join('')}</table>`; }
    return `<li class="ask"><i>?</i><div><b>${NEED[n.k]}</b>${w}</div></li>`;
  }
  function body(uid) {
    const u = U(uid);
    switch (u.status) {
      case 'locked': return `<p class="muted">Unlocks after ${UNITS[Object.keys(UNITS).find(k => UNITS[k].next === uid)].name} sign-off.</p><button class="link" data-act="early" data-u="${uid}">Start early</button>`;
      case 'ready': return `<button class="btn primary" data-act="check" data-u="${uid}">Start</button>`;
      case 'checking': return `<p class="muted"><span class="spin"></span>Checking the workbook…</p>`;
      case 'ctx': return `<ul class="needs">${u.needs.map(n => needRow(uid, n)).join('')}</ul><div class="acts"><button class="btn primary" data-act="ctx" data-u="${uid}">Continue</button></div>`;
      case 'plan': {
        const sizes = procsOf(uid).includes('samp') ? `<div class="sizes"><b>Sample sizes</b>${D.AREAS.map(a => `<label>${a.name}<input type="number" min="1" max="60" data-size="${a.id}" value="${ST.sizes[a.id]}"></label>`).join('')}</div>` : '';
        const steps = procsOf(uid).map(p => `<div class="grp">${PROCS[p].ref} ${PROCS[p].name}</div>${PROCS[p].build(dummyCtx()).map(s => `<li>${esc(s.step)}</li>`).join('')}`).join('');
        return `${u.err ? `<p class="err">Excel error: ${esc(u.err)}</p>` : ''}${sizes}<ol class="steps">${steps}</ol><div class="acts"><button class="btn primary" data-act="run" data-u="${uid}">Approve & run</button><button class="btn ghost" data-act="back" data-u="${uid}">Back</button></div>`;
      }
      case 'running': {
        const all = procsOf(uid).flatMap(p => PROCS[p].build(dummyCtx()).map(s => s.step));
        return `<div class="bar"><i style="width:${Math.round(u.done.length / all.length * 100)}%"></i></div><ol class="steps live">${all.map(s => `<li class="${u.done.includes(s) ? 'd' : s === u.cur ? 'a' : ''}">${esc(s)}</li>`).join('')}</ol>`;
      }
      case 'review':
      case 'done': {
        const sheets = procsOf(uid).flatMap(p => PROCS[p].sheets);
        return `${(u.results || []).map(r => `<p class="res">${esc(r)}</p>`).join('')}<div class="chips">${sheets.map(s => `<button class="chip" data-act="goto" data-s="${esc(s)}">${esc(s)}</button>`).join('')}</div>
          ${u.status === 'review' ? `<div class="acts"><button class="btn primary" data-act="signoff" data-u="${uid}">Sign off</button><button class="btn ghost" data-act="check" data-u="${uid}">Re-run</button></div>` : `<p class="signed">✓ Signed off ${esc(u.signed || '')}</p><button class="link" data-act="check" data-u="${uid}">Re-run</button>`}`;
      }
    }
    return '';
  }
  function card(uid, n) {
    const u = U(uid);
    const sub = procsOf(uid).map(p => PROCS[p].ref).join(' · ');
    return `<section class="unit ${u.status}"><header><span class="n">${n}</span><div><b>${nameOf(uid)}</b><small>${sub}</small></div>${pill(u.status)}</header><div class="ub">${body(uid)}</div></section>`;
  }
  function viewSetup() {
    const s = ST.setup || {client:'Harrow Precision Components (Pty) Ltd', ye:'2026-09-30', framework:'IFRS for SMEs', entity:'earnings', reviewer:'LD', users:'cfo.vdm, admin', tbSheet:'TB', glSheet:'GL'};
    return `<div class="setup"><h2>Engagement</h2>
      <label>Client<input id="s-client" value="${esc(s.client)}"></label>
      <div class="row"><label>Year end<input id="s-ye" type="date" value="${esc(s.ye)}"></label><label>Framework<select id="s-fw">${['IFRS for SMEs','IFRS'].map(f => `<option ${f === s.framework ? 'selected' : ''}>${f}</option>`).join('')}</select></label></div>
      <label>Materiality basis<select id="s-entity">${Object.entries(WP.BASIS).map(([k, v]) => `<option value="${k}" ${k === s.entity ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
      <div class="row"><label>Reviewer initials<input id="s-rev" value="${esc(s.reviewer)}" maxlength="4"></label><label>Senior users<input id="s-users" value="${esc(s.users)}"></label></div>
      <h3>Ledger data</h3>
      <div class="row"><label>GL sheet<input id="s-gl" value="${esc(s.glSheet)}"></label><label>TB sheet<input id="s-tb" value="${esc(s.tbSheet)}"></label></div>
      <button class="btn" data-act="sample">Create sample GL & TB</button>
      <div class="acts"><button class="btn primary" data-act="save-setup">Save</button>${ST.setup ? '<button class="btn ghost" data-act="close-setup">Cancel</button>' : ''}</div></div>`;
  }
  function viewAudit() {
    if (!ST.setup) return `<div class="empty"><p>Set up the engagement to start.</p><button class="btn primary" data-act="open-setup">Set up</button></div>`;
    return Object.keys(UNITS).map(id => card(id, UNITS[id].n)).join('');
  }
  function viewProcs() {
    if (!ST.setup) return viewAudit();
    return `<p class="muted intro">Run any process on its own. Pipe checks the workbook and asks for what's missing.</p>` + Object.entries(PROCS).map(([p, d]) => {
      const uid = 'p:' + p, u = U(uid), open = ST.open === p;
      return `<section class="unit proc ${u.status} ${open ? 'open' : ''}"><header data-act="toggle" data-p="${p}"><span class="ref">${d.ref}</span><div><b>${d.name}</b></div>${pill(u.status)}</header>${open ? `<div class="ub">${body(uid)}</div>` : ''}</section>`;
    }).join('');
  }
  let chatSuggest = [];
  function viewChat() {
    return `<div class="chat"><div class="msgs" id="msgs">${ST.chat.map(m => `<div class="m ${m.who}">${md(m.text)}</div>`).join('')}</div>
      <div class="sugg">${chatSuggest.map(s => `<button data-act="ask" data-q="${esc(s)}">${esc(s)}</button>`).join('')}</div>
      <form id="chat-form"><input id="chat-in" placeholder="Ask Pipe…" autocomplete="off"><button class="btn primary" type="submit">Send</button></form></div>`;
  }
  function draw() {
    $('#client').textContent = ST.setup ? `${ST.setup.client.replace(/ \(Pty\) Ltd$/, '')} · ${ST.setup.ye}` : 'No engagement';
    document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === ST.tab && !ST.view));
    const scroll = $('#main').scrollTop;
    $('#main').innerHTML = ST.view === 'setup' ? viewSetup() : ST.tab === 'audit' ? viewAudit() : ST.tab === 'procs' ? viewProcs() : viewChat();
    $('#main').scrollTop = scroll;
    if (ST.tab === 'chat' && !ST.view) { const m = $('#msgs'); m.scrollTop = 1e6; }
    save();
  }
  function busy(on) { $('#app').classList.toggle('busy', on); }
  function toast(t) { const el = $('#toast'); el.textContent = t; el.classList.add('show'); clearTimeout(toast.t); toast.t = setTimeout(() => el.classList.remove('show'), 2400); }

  /* ---------- Chat ---------- */
  function md(t) { return esc(t).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').split('\n').map(l => l.startsWith('- ') ? `<div class="li">${l.slice(2)}</div>` : `<p>${l}</p>`).join(''); }
  function say(text) {
    ST.chat.push({who:'ai', text}); ST.chat = ST.chat.slice(-40);
    if (ST.tab === 'chat') draw(); else { $('.tabs [data-tab="chat"]').classList.add('dot'); save(); }
  }
  async function refreshSuggest() {
    let sheet = '';
    try { sheet = await XL.activeSheet() || ''; } catch (e) {}
    const ref = sheet.split(' ')[0];
    chatSuggest = {A1:['Why profit before tax?','What can I change here?'], A2:['Why is revenue significant?','How is the rating worked out?'], A3:['Tell me about JE-24117','How is the score worked out?'], B1:['How were sample sizes set?','Why MUS for revenue?'], S1:['Do we need an adjustment?']}[ref]
      || (/^[DEFCJ]\d\.\d$/.test(ref) ? ['Summarise the exceptions','How is projection calculated?'] : ['What can you do?','Run one process on its own?']);
    chatSuggest.ctx = sheet;
  }
  function answer(q) {
    const t = q.toLowerCase(), has = (...w) => w.some(x => t.includes(x));
    if (has('on its own','one process','single')) return 'Yes. Open **Processes** and pick one. I check the workbook for what it needs (for example A1 for materiality) and ask only for what is missing. Anything you give me goes into a shaded cell marked as auditor-provided.';
    if (has('change','edit','shaded')) return 'Shaded cells are judgements: rates, likelihood and magnitude, sample sizes, test results. Change any of them and the working paper recalculates. Unshaded figures are formulas linked to the TB, GL or other papers.';
    if (has('profit before tax','materiality','basis')) return 'Profit before tax is used because Harrow is profit-oriented and its lenders focus on earnings. A1 takes the benchmarks from the TB with SUMIFS, applies 5% and rounds down to R 10,000 (J17). PM is 75% (J19) and clearly trivial 5% (J21).';
    if (has('24117','journal','score')) return 'Each manual journal scores 1 point per criterion: weekend, round amount, near year end, post-close, senior user and unusual pairing. All but pairing are formulas in A3. **JE-24117** scores 4: R 450,000 of sales posted on Sunday 27 September and reversed in P13 by JE-24203.';
    if (has('rating','revenue','significant')) return 'In A2 the rating is a formula on likelihood and magnitude: High + High is Significant, journals are presumed significant. Revenue is High/High because September is well above the other months (F22) and JE-24117 moved revenue across year end.';
    if (has('sample size','mus','sampling')) return 'Sizes are judgements in B1 column G, checked against the firm minimums (C37:D40) in column L. Revenue uses MUS because the risk is overstatement. The interval is population ÷ sample size (column H).';
    if (has('exception','summar')) return 'Detail testing found:\n- **REV-07** cut-off: R 612,500 delivered 2 Oct\n- **REV-18** price difference R 12,400\n- **INV-04** count shortfall R 41,600\n- **PAY-09** unrecorded liability R 186,000\n- **JE-24117** no approval\n- **PPE-06** no support yet';
    if (has('projection','projected')) return 'For revenue and inventory, projected = known misstatement × population ÷ sample value (L15). Other areas use known misstatement only.';
    if (has('adjustment','opinion')) return 'S1 compares projected misstatement with materiality. Revenue cut-off pushes the total above overall materiality, so propose the REV-07 adjustment before signing.';
    if (has('what can you','help')) return 'I set up the audit file, then write each working paper into the workbook: A1 materiality, A2 risks, A3 journals, B1 sampling, one detail testing sheet per area and S1 misstatements. Calculations are live formulas, and you approve each plan and sign off each result.';
    return 'I only have scripted answers in this demo. Try one of the suggestions.';
  }
  async function ask(q) {
    if (!q.trim()) return;
    ST.chat.push({who:'me', text:q}); draw();
    await sleep(500);
    say(answer(q));
  }

  /* ---------- Preview workbook (outside Excel) ---------- */
  function fmtCell(c) {
    const v = c.v, nf = c.f.nf || '';
    if (typeof v === 'string' && v.startsWith('=')) return `<span class="fx">${esc(v.length > 60 ? v.slice(0, 57) + '…' : v)}</span>`;
    if (typeof v === 'number') {
      if (nf.includes('mmm')) return D.fmtDate(v);
      if (nf.includes('%')) return (v * 100).toFixed(nf.includes('.0') ? 1 : 0) + '%';
      if (nf.includes('#,##0')) return v.toLocaleString('en-US', {minimumFractionDigits:nf.includes('.00') ? 2 : 0, maximumFractionDigits:2});
    }
    return esc(v);
  }
  function drawPreview() {
    const b = XL.book, el = $('#preview');
    if (!b.order.length) { el.innerHTML = '<div class="pv-empty">Workbook preview. Sheets Pipe writes appear here.</div>'; return; }
    const s = b.sheets[b.active] || b.sheets[b.order[0]];
    let mr = 1, mc = 1;
    Object.keys(s.cells).forEach(k => { const x = XLUtil.parse(k); mr = Math.max(mr, x.r1); mc = Math.max(mc, x.c1); });
    let h = '<table class="grid"><tr><th></th>';
    for (let c = 1; c <= mc; c++) h += `<th style="width:${(s.cols[XLUtil.colName(c)] || 9) * 7}px">${XLUtil.colName(c)}</th>`;
    h += '</tr>';
    for (let r = 1; r <= Math.min(mr, 600); r++) {
      h += `<tr${s.rows[r] ? ` style="height:${s.rows[r] * 1.3}px"` : ''}><th>${r}</th>`;
      for (let c = 1; c <= mc; c++) {
        const cell = s.cells[XLUtil.colName(c) + r];
        if (!cell) { h += '<td></td>'; continue; }
        const f = cell.f, st = [f.b ? 'font-weight:700' : '', f.i ? 'font-style:italic' : '', f.c ? `color:${f.c}` : '', f.fill ? `background:${f.fill}` : '', f.h ? `text-align:${f.h.toLowerCase()}` : '', f.bd ? 'border-bottom:1px solid #9FB3C2' : '', f.w ? 'white-space:normal' : ''].filter(Boolean).join(';');
        h += `<td style="${st}">${fmtCell(cell)}</td>`;
      }
      h += '</tr>';
    }
    el.innerHTML = `<div class="pv-tabs">${b.order.map(n => `<button class="${n === b.active ? 'on' : ''}" data-act="pv" data-s="${esc(n)}" style="${b.sheets[n].tab ? `border-top:3px solid ${b.sheets[n].tab}` : ''}">${esc(n)}</button>`).join('')}</div><div class="pv-grid">${h}</table></div>`;
  }

  /* ---------- Events ---------- */
  document.addEventListener('click', async e => {
    const el = e.target.closest('[data-act]');
    if (!el || el.disabled) return;
    const uid = el.dataset.u;
    switch (el.dataset.act) {
      case 'open-setup': ST.view = 'setup'; draw(); break;
      case 'close-setup': ST.view = null; draw(); break;
      case 'save-setup': saveSetup(); break;
      case 'sample': sampleData(); break;
      case 'check': check(uid); break;
      case 'early': U(uid).status = 'ready'; check(uid); break;
      case 'ctx': confirmCtx(uid); break;
      case 'run': run(uid); break;
      case 'back': U(uid).status = 'ready'; draw(); break;
      case 'signoff': signOff(uid); break;
      case 'docs': requestDocs(); break;
      case 'goto': try { await XL.activate(el.dataset.s); } catch (err) { toast(err.message); } break;
      case 'toggle': ST.open = ST.open === el.dataset.p ? null : el.dataset.p; draw(); break;
      case 'ask': ask(el.dataset.q); break;
      case 'pv': XL.book.active = el.dataset.s; drawPreview(); break;
      case 'tab': ST.tab = el.dataset.tab; ST.view = null; if (ST.tab === 'chat') { el.classList.remove('dot'); await refreshSuggest(); } draw(); break;
    }
  });
  document.addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.size) { ST.sizes[t.dataset.size] = Math.max(1, Math.min(60, parseInt(t.value) || 1)); draw(); }
  });
  document.addEventListener('submit', e => {
    if (e.target.id !== 'chat-form') return;
    e.preventDefault(); const i = $('#chat-in'); const q = i.value; i.value = ''; ask(q);
  });

  /* ---------- Boot ---------- */
  async function boot(info) {
    inExcel = !!(info && info.host === Office.HostType.Excel);
    XL = makeXL(inExcel);
    document.body.classList.toggle('preview', !inExcel);
    if (!inExcel) { XL.onChange(drawPreview); drawPreview(); window.PipeXL = XL; }
    load();
    if (!ST.setup && inExcel && await XL.exists('Index')) {
      try {
        const v = await XL.read('Index', 'D4:D10');
        const ye = D.fromSerial(v[1][0]);
        ST.setup = {client:v[0][0], ye:`${ye.y}-${String(ye.m).padStart(2, '0')}-${String(ye.d).padStart(2, '0')}`, framework:v[2][0], reviewer:v[4][0], entity:Object.keys(WP.BASIS).find(k => WP.BASIS[k] === v[5][0]) || 'earnings', users:v[6][0], tbSheet:'TB', glSheet:'GL'};
      } catch (e) {}
    }
    if (!ST.setup) ST.view = 'setup';
    await refreshSuggest();
    draw();
  }
  if (window.Office && Office.onReady) Office.onReady(boot); else boot(null);
})();
