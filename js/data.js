/* Demo data for the Harrow Precision engagement. Everything here is fictional. */
(function (g) {
  const serial = (y, m, d) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  const fromSerial = s => { const d = new Date(Date.UTC(1899, 11, 30) + s * 86400000); return {y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), wd: d.getUTCDay()}; };
  const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const fmtDate = s => { const x = fromSerial(s); return `${x.d} ${MON[x.m - 1]} ${x.y}`; };
  function rng(seed) { return () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; }; }
  function hash(s) { let h = 2166136261; for (const c of s) { h ^= c.charCodeAt(0); h = Math.imul(h, 16777619); } return (h >>> 0) / 4294967296; }

  const YE = serial(2026, 9, 30);

  // Account, name, balance (Dr +, Cr −), FS line, class
  const TB = [
    [1000, 'Bank – current account', 4820000, 'Cash and cash equivalents', 'Asset'],
    [1100, 'Trade receivables control', 31600000, 'Trade and other receivables', 'Asset'],
    [1300, 'Raw materials', 12400000, 'Inventories', 'Asset'],
    [1310, 'Work in progress', 9100000, 'Inventories', 'Asset'],
    [1320, 'Finished goods', 19600000, 'Inventories', 'Asset'],
    [1390, 'Obsolescence provision', -2200000, 'Inventories', 'Asset'],
    [1500, 'Plant & machinery – cost', 64300000, 'Property, plant and equipment', 'Asset'],
    [1510, 'Capital work in progress', 12300000, 'Property, plant and equipment', 'Asset'],
    [1590, 'Accumulated depreciation', -21600000, 'Property, plant and equipment', 'Asset'],
    [2100, 'Trade payables control', -17900000, 'Trade and other payables', 'Liability'],
    [2150, 'Accrued expenses', -3800000, 'Trade and other payables', 'Liability'],
    [2200, 'VAT payable', -2100000, 'Trade and other payables', 'Liability'],
    [2500, 'Term loan', -28000000, 'Borrowings', 'Liability'],
    [3000, 'Share capital', -5000000, 'Equity', 'Equity'],
    [3100, 'Retained earnings – opening', -58650000, 'Equity', 'Equity'],
    [4000, 'Sales – domestic', -142700000, 'Revenue', 'Revenue'],
    [4010, 'Sales – export', -46100000, 'Revenue', 'Revenue'],
    [4900, 'Returns & rebates', 2400000, 'Revenue', 'Revenue'],
    [5000, 'Cost of sales', 139200000, 'Cost of sales', 'Expense'],
    [6000, 'Salaries & wages', 18400000, 'Operating expenses', 'Expense'],
    [6100, 'Depreciation', 6900000, 'Operating expenses', 'Expense'],
    [6200, 'Administration & other', 4830000, 'Operating expenses', 'Expense'],
    [7000, 'Finance costs', 2200000, 'Finance costs', 'Expense']
  ];
  const ACCT = Object.fromEntries(TB.map(r => [r[0], r[1]]));

  // GL accounts mapped to audit areas for sampling
  const MAPPING = [
    [1000, 'Not sampled', 'Covered by bank confirmation'],
    [1100, 'Not sampled', 'Covered by debtor confirmations'],
    [1300, 'Inventory', ''], [1310, 'Inventory', ''], [1320, 'Inventory', ''], [1390, 'Inventory', ''],
    [1500, 'Not sampled', 'Opening balance; additions sit in 1510'],
    [1510, 'PPE', ''],
    [2100, 'Payables', ''], [2150, 'Payables', ''],
    [4000, 'Revenue', ''], [4010, 'Revenue', ''], [4900, 'Revenue', '']
  ];

  const MONTHS = ['Oct','Nov','Dec','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep'];
  const REV_CY = [14.2,14.8,13.1,13.9,14.6,15.3,15.0,15.8,15.4,16.1,16.6,21.6].map(v => v * 1e6);
  const REV_PY = [13.6,14.0,12.5,13.2,13.8,14.4,14.1,14.6,14.3,14.9,15.2,15.6].map(v => v * 1e6);

  // Likelihood / magnitude are Pipe's judgements; the rating is a formula in A2
  const RISKS = [
    {area:'Revenue', as:'Occurrence, cut-off', l:'High', m:'High', resp:'MUS tests of detail plus extended cut-off testing around year end.', src:'ISA 240 ¶26'},
    {area:'Journals', as:'Management override', l:'Presumed', m:'High', resp:'Targeted testing of flagged journals and review of post-close entries.', src:'Methodology §6.3'},
    {area:'Inventory', as:'Existence, valuation', l:'Medium', m:'High', resp:'Stratified test counts, key items and an NRV review of slow-moving SKUs.', src:'Risk library MFG-04'},
    {area:'PPE', as:'Existence, valuation', l:'Medium', m:'Medium', resp:'Vouch additions above PM plus a random top-up.', src:'FY2025 WP C2'},
    {area:'Payables', as:'Completeness', l:'Medium', m:'Medium', resp:'Search for unrecorded liabilities using post year-end payments.', src:'Risk library MFG-11'},
    {area:'Receivables', as:'Existence, valuation', l:'Medium', m:'Medium', resp:'Debtor confirmations and subsequent receipts.', src:'Methodology §7.2'},
    {area:'Cash', as:'Existence', l:'Low', m:'Low', resp:'Bank confirmations.', src:'Methodology §7.1'},
    {area:'Payroll', as:'Occurrence', l:'Low', m:'Low', resp:'Analytical review only.', src:'FY2025 WP P1'}
  ];
  const RISK_WHY = {
    Revenue: '="September revenue is "&TEXT(F22,"0%")&" above the average of the other months. JE-24117 (R 450,000) is a manual sales accrual posted on Sunday 27 September and reversed after year end."',
    Journals: 'Presumed significant under ISA 240. Several manual journals meet three or more risk criteria (see A3), including entries by the CFO around year end.',
    Inventory: 'Inventory days rose from 61 to 74. JE-23980 released R 1.2m of obsolescence provision on 30 September with no supporting analysis.',
    PPE: 'R 12.3m of additions sits in capital work in progress, mostly the new CNC line. Costs may be capitalised before the asset is ready for use.',
    Payables: 'Accrued expenses fell 22% while purchases rose 9%, which points to possible unrecorded liabilities.',
    Receivables: 'Debtor days are stable at 58, but two large customers are over 120 days.',
    Cash: 'Four bank accounts, reconciled monthly with no long-outstanding items.',
    Payroll: 'Headcount and payroll cost are in line with the prior year.'
  };

  // Sampling areas. a2Row is the risk row in A2; b1Row is the plan row in B1.
  const AREAS = [
    {id:'REV', name:'Revenue', sheet:'D3.1 Revenue', ref:'D3.1', full:'Revenue – occurrence & cut-off', assert:'Occurrence, cut-off', method:'Monetary unit sampling', size:25, popV:186.4e6, project:true, a2Row:29, b1Row:29, doc:'Sales invoice + delivery note', checks:'invoice exists, amount and customer agree, goods delivered within the period'},
    {id:'INV', name:'Inventory', sheet:'E2.1 Inventory', ref:'E2.1', full:'Inventory – existence & valuation', assert:'Existence, valuation', method:'Stratified + key items', size:15, popV:38.9e6, project:true, a2Row:31, b1Row:30, doc:'Count sheet + standard cost card', checks:'counted quantity agrees with the system and unit cost agrees with the cost card'},
    {id:'PAY', name:'Payables', sheet:'F4.1 Payables', ref:'F4.1', full:'Trade payables – completeness', assert:'Completeness', method:'Unrecorded liabilities search', size:12, popV:21.7e6, project:false, a2Row:33, b1Row:31, doc:'Supplier invoice + GRN', checks:'the service or delivery date is compared with year end and the liability is recorded in the right period'},
    {id:'PPE', name:'PPE', sheet:'C2.1 PPE', ref:'C2.1', full:'PPE additions – existence & valuation', assert:'Existence, valuation', method:'Key items + random', size:8, popV:12.3e6, project:false, a2Row:32, b1Row:32, doc:'Supplier invoice + asset register', checks:'the invoice supports cost and the asset exists and is correctly classified'},
    {id:'JE', name:'Journals', sheet:'J1.1 Journals', ref:'J1.1', full:'Journals – management override', assert:'Occurrence, authorisation', method:'Targeted (risk criteria)', size:10, popV:0, project:false, a2Row:30, b1Row:33, doc:'Journal voucher + approval', checks:'there is a business rationale, supporting evidence and independent approval'}
  ];
  const RANGE = {REV:[60000,1400000], INV:[50000,900000], PAY:[40000,600000], PPE:[90000,2400000]};
  const PARTIES = {
    REV:['Pretoria Steelworks','Mabena Mining Supplies','Karoo Agri Equipment','Coastal Pump Systems','Vaal Automotive','Lekker Hydraulics','Bushveld Conveyors','Umhlanga Marine'],
    INV:['Raw – 316L bar stock','FG – Pump housing PH-220','WIP – Gear blank GB-40','FG – Flange set F-150','Raw – Cast iron billets','FG – Valve body VB-12','FG – Impeller IM-75'],
    PAY:['Sasko Tooling','Ngwenya Freight','City Power','Metalco Alloys','Cape Abrasives','Sibanye Industrial Gas'],
    PPE:['DMG CNC lathe line 3','Haas 5-axis mill','Toyota 3t forklift','Compressor upgrade','Factory racking','Quality lab CMM','Extraction system']
  };

  // Manual journals analysed in A3. judg = Pipe's judgement on the account pairing.
  const KEY_JOURNALS = [
    {je:'JE-24117', date:serial(2026,9,27), user:'t.nkosi', desc:'Sales accrual – Pretoria Steelworks', pair:'Dr 1100 Receivables / Cr 4000 Sales', dr:1100, cr:4000, amt:450000, judg:1},
    {je:'JE-23980', date:serial(2026,9,30), user:'cfo.vdm', desc:'Inventory provision release', pair:'Dr 1390 Provision / Cr 5000 Cost of sales', dr:1390, cr:5000, amt:1200000, judg:0},
    {je:'JE-24203', date:serial(2026,10,3), user:'cfo.vdm', desc:'Reverse sales accrual JE-24117', pair:'Dr 4000 Sales / Cr 1100 Receivables', dr:4000, cr:1100, amt:450000, judg:0},
    {je:'JE-21877', date:serial(2026,7,14), user:'admin', desc:'Repairs accrual (posted 22:41)', pair:'Dr 6200 Admin / Cr 2150 Accruals', dr:6200, cr:2150, amt:85000, judg:1},
    {je:'JE-22510', date:serial(2026,8,31), user:'j.pillay', desc:'Reclass repairs to capital WIP', pair:'Dr 1510 Capital WIP / Cr 6200 Admin', dr:1510, cr:6200, amt:640000, judg:1},
    {je:'JE-19034', date:serial(2026,3,2), user:'t.nkosi', desc:'Rebate true-up', pair:'Dr 4900 Rebates / Cr 2100 Payables', dr:4900, cr:2100, amt:100000, judg:1}
  ];
  const ROUTINE = [
    ['Payroll accrual', 6000, 2150], ['FX revaluation', 1100, 4010], ['Freight accrual', 5000, 2150], ['Depreciation adjustment', 6100, 1590],
    ['Intercompany recharge', 6200, 2100], ['Prepayment release', 6200, 1100], ['Bank charges correction', 7000, 1000], ['Accrual reversal', 2150, 6200]
  ];
  function journals() {
    const r = rng(7341), out = KEY_JOURNALS.map(j => ({...j}));
    for (let i = 0; i < 24; i++) {
      const [desc, dr, cr] = ROUTINE[i % ROUTINE.length];
      let s = serial(2025, 10, 1) + Math.floor(r() * 355);
      if (s >= YE - 4) s -= 9;
      const wd = fromSerial(s).wd;
      if (wd === 0) s += 1; if (wd === 6) s += 2;
      const amt = Math.round((8000 + r() * 260000) * 100) / 100 + 0.37;
      out.push({je:'JE-' + (19100 + Math.floor(r() * 4900)), date:s, user:i % 3 ? 'j.pillay' : 'm.botha', desc, pair:`Dr ${dr} ${ACCT[dr].split(' ')[0]} / Cr ${cr} ${ACCT[cr].split(' ')[0]}`, dr, cr, amt, judg:0});
    }
    return out;
  }

  // General ledger extract: system sales and purchases plus every manual journal
  function ledger() {
    const r = rng(90210), lines = [];
    let inv = 41000, pur = 7700;
    for (let mi = 0; mi < 12; mi++) {
      const y = mi < 3 ? 2025 : 2026, m = mi < 3 ? 10 + mi : mi - 2;
      for (let k = 0; k < 12; k++) {
        const d = serial(y, m, 1 + Math.floor(r() * 27)), amt = Math.round(40000 + r() * 1300000), c = PARTIES.REV[Math.floor(r() * 8)];
        const no = 'SI-' + (inv++);
        lines.push([d, no, 1100, ACCT[1100], 'Invoice – ' + c, amt, '', 'System', 'SYSTEM']);
        lines.push([d, no, k % 4 === 0 ? 4010 : 4000, ACCT[k % 4 === 0 ? 4010 : 4000], 'Invoice – ' + c, '', amt, 'System', 'SYSTEM']);
      }
      for (let k = 0; k < 6; k++) {
        const d = serial(y, m, 1 + Math.floor(r() * 27)), amt = Math.round(20000 + r() * 700000), s = PARTIES.PAY[Math.floor(r() * 6)];
        const no = 'PI-' + (pur++);
        lines.push([d, no, 5000, ACCT[5000], 'Purchase – ' + s, amt, '', 'System', 'SYSTEM']);
        lines.push([d, no, 2100, ACCT[2100], 'Purchase – ' + s, '', amt, 'System', 'SYSTEM']);
      }
    }
    journals().forEach(j => {
      lines.push([j.date, j.je, j.dr, ACCT[j.dr], j.desc, j.amt, '', 'Manual', j.user]);
      lines.push([j.date, j.je, j.cr, ACCT[j.cr], j.desc, '', j.amt, 'Manual', j.user]);
    });
    return lines.sort((a, b) => a[0] - b[0]);
  }

  function score(j) {
    const wd = fromSerial(j.date).wd;
    return (wd === 0 || wd === 6 ? 1 : 0) + (j.amt % 10000 === 0 ? 1 : 0) + (j.date <= YE && j.date >= YE - 3 ? 1 : 0) + (j.date > YE ? 1 : 0) + (/cfo\.vdm|admin/.test(j.user) ? 1 : 0) + j.judg;
  }

  // Samples, with the exceptions Pipe "finds" in detail testing
  function samples(sizes) {
    const r = rng(20261102), out = [];
    AREAS.forEach(a => {
      const n = sizes[a.id] || 0, list = [];
      if (a.id === 'JE') {
        journals().map(j => ({j, s:score(j)})).sort((x, y) => y.s - x.s || y.j.amt - x.j.amt).slice(0, n).forEach(({j, s}, i) => {
          list.push({id:'JE-' + String(i + 1).padStart(2, '0'), area:'JE', ref:j.je, date:j.date, party:j.desc, amt:j.amt, why:s >= 3 ? `${s} criteria met` : s ? `${s} criterion met` : 'Random top-up', user:j.user});
        });
      } else {
        for (let i = 1; i <= n; i++) {
          const [lo, hi] = RANGE[a.id];
          const amt = Math.round((lo + Math.pow(r(), 1.6) * (hi - lo)) / 10) * 10;
          const party = PARTIES[a.id][Math.floor(r() * PARTIES[a.id].length)];
          const base = Math.floor(r() * 9000);
          let date;
          if (a.id === 'PAY') date = serial(2026, 10, 1) + Math.floor(r() * 58);
          else if (a.id === 'INV') date = YE;
          else date = serial(2025, 10, 1) + Math.floor(r() * 364);
          const ref = {REV:'SI-' + (40000 + base), INV:'SKU-' + (1000 + base % 1200), PAY:'PMT-' + (7000 + base), PPE:'FA-' + (300 + base % 90)}[a.id];
          const why = {REV:'MUS interval hit', INV:i <= 3 ? 'Key item > PM' : 'Stratum ' + (i % 2 + 1), PAY:'Post year-end payment', PPE:i <= 3 ? 'Key item > PM' : 'Random'}[a.id];
          list.push({id:a.id + '-' + String(i).padStart(2, '0'), area:a.id, ref, date, party, amt, why});
        }
      }
      const at = k => list[Math.min(k, list.length - 1)];
      if (!list.length) return;
      if (a.id === 'REV') {
        Object.assign(at(6), {party:'Pretoria Steelworks', amt:612500, date:serial(2026, 9, 28), x:{type:'cutoff', docDate:serial(2026, 10, 2), note:'Delivered 2 Oct 2026; revenue recorded 28 Sep 2026. Belongs to FY2027.'}});
        if (list.length >= 10) { const s = at(17); if (!s.x) s.x = {type:'amount', docAmt:s.amt - 12400, note:'Invoice R 12,400 below GL. August price list not applied.'}; }
      }
      if (a.id === 'INV') Object.assign(at(3), {party:'FG – Pump housing PH-220', amt:275600, x:{type:'count', docAmt:275600 - 41600, note:'Counted 180 units vs system 212. 32 units × R 1,300 standard cost.'}});
      if (a.id === 'PAY') Object.assign(at(8), {party:'Ngwenya Freight', amt:186000, date:serial(2026, 10, 14), x:{type:'unrecorded', docDate:serial(2026, 9, 22), note:'Invoice dated 22 Sep 2026 for September freight, paid 14 Oct 2026, not accrued.'}});
      if (a.id === 'PPE') Object.assign(at(5), {x:{type:'missing', note:'No supplier invoice found in 224 documents. Requested from client.'}});
      if (a.id === 'JE') { const s = list.find(x => x.ref === 'JE-24117'); if (s) s.x = {type:'approval', note:'No approver signature. Posted on a Sunday by the preparer, reversed in P13 (JE-24203). Quantified in D3.1.'}; }
      out.push(...list);
    });
    return out;
  }
  function docNo(s) {
    const n = Math.floor(hash(s.id + 'd') * 9000) + 1000;
    return {REV:'INV-' + (88000 + n % 4000), INV:'CS-30S-' + String(n % 120).padStart(3, '0'), PAY:'SUP-' + n, PPE:'FA-INV-' + n, JE:'JV-' + s.ref.slice(3)}[s.area];
  }
  // Test result for one sample: doc fields, result and how misstatement is calculated
  function test(s) {
    const x = s.x || {};
    if (x.type === 'missing') return {doc:'', docDate:'', docAmt:'', result:'No support', mis:'zero', note:x.note};
    return {
      doc:docNo(s), docDate:x.docDate || s.date, docAmt:x.docAmt != null ? x.docAmt : s.amt,
      result:s.x ? 'Exception' : 'Agreed',
      mis:{cutoff:'full', unrecorded:'full', amount:'diff', count:'diff', approval:'zero'}[x.type] || 'none',
      note:x.note || ''
    };
  }
  function evalArea(a, list) {
    const sv = list.reduce((t, s) => t + s.amt, 0);
    let known = 0, exc = 0, missing = 0;
    list.forEach(s => {
      const t = test(s);
      if (t.result === 'Exception') exc++;
      if (t.result === 'No support') missing++;
      if (t.mis === 'full') known += s.amt;
      if (t.mis === 'diff') known += s.amt - t.docAmt;
    });
    return {n:list.length, sv, known, exc, missing, projected:a.project && sv ? known * a.popV / sv : known};
  }

  g.PipeData = {serial, fromSerial, fmtDate, hash, YE, TB, ACCT, MAPPING, MONTHS, REV_CY, REV_PY, RISKS, RISK_WHY, AREAS, journals, ledger, score, samples, test, evalArea, docNo};
})(window);
