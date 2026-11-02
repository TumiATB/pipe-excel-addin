/* Working-paper builders. Each process returns sections; each section writes part of
   a sheet. Calculations are live formulas; Pipe's judgements go in shaded input cells. */
(function (g) {
  const D = g.PipeData;
  const ACC = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
  const DATE = 'd-mmm-yy', PCT = '0.0%', PCT0 = '0%';
  const INPUT = '#FFF6D6', HEAD = '#E3EEF5', TITLE = '#1F5469', NOTE = '#6B7B88', GOOD = '#2F8F5B', BAD = '#C62828', BAD_FILL = '#FBE4E4', WARN = '#A86B0C', WARN_FILL = '#FBEFD9';

  const ORDER = ['Index','A1 Materiality','A2 Risk assessment','A3 Journals','B1 Sampling','D3.1 Revenue','E2.1 Inventory','F4.1 Payables','C2.1 PPE','J1.1 Journals','S1 Misstatements','GL','TB'];
  const posFor = me => names => {
    const others = names.filter(n => n !== me);
    const after = others.find(n => ORDER.indexOf(n) > ORDER.indexOf(me));
    return after ? others.indexOf(after) : others.length;
  };
  const WPS = [
    ['A1','A1 Materiality','Materiality','mat'],
    ['A2','A2 Risk assessment','Risk assessment','risk'],
    ['A3','A3 Journals','Journal testing','je'],
    ['B1','B1 Sampling','Sampling','samp'],
    ...D.AREAS.map(a => [a.ref, a.sheet, 'Detail testing', 'test']),
    ['S1','S1 Misstatements','Misstatements','sum']
  ];
  const BASIS = {earnings:'Profit before tax', revenue:'Revenue', assets:'Total assets'};
  const q = n => `'${n.replace(/'/g, "''")}'`;
  const R = n => 'R ' + Math.round(n).toLocaleString('en-US');
  const today = () => { const d = new Date(); return D.serial(d.getFullYear(), d.getMonth() + 1, d.getDate()); };

  function header(o, ref, title, last = 'J') {
    o.v('J1', 'Ref# ' + ref); o.f('J1', {b:true, h:'Right', c:TITLE});
    o.v('B2', 'Client Name'); o.v('D2', '=Index!$D$4'); o.v('H2', 'Prepared by'); o.v('I2', '=Index!$D$7'); o.v('J2', today());
    o.v('B3', 'Year End'); o.v('D3', '=Index!$D$5'); o.v('H3', 'Reviewed by');
    o.f('B2:B3', {b:true}); o.f('H2:H3', {b:true}); o.f('D3', {nf:DATE, h:'Left'}); o.f('J2:J3', {nf:DATE, h:'Right'});
    o.f(`B3:${last}3`, {bd:'bottom'});
    o.v('B5', title); o.f('B5', {b:true, sz:13, c:TITLE});
  }
  const heading = (o, a, t) => { o.v(a, t); o.f(a, {b:true, c:TITLE}); };
  const note = (o, a, t) => { o.v(a, t); o.f(a, {i:true, c:NOTE}); };
  const thead = (o, a, cols) => { o.v(a, [cols]); const x = g.XLUtil.parse(a); const end = g.XLUtil.colName(x.c1 + cols.length - 1) + x.r1; o.f(`${a}:${end}`, {b:true, fill:HEAD, bd:'bottom', w:true, va:'Bottom'}); };
  const para = (o, a, range, t, h) => { o.v(a, t); o.f(range, {m:true, w:true, va:'Top'}); if (h) { const x = g.XLUtil.parse(range); const rows = {}; for (let r = x.r1; r <= x.r2; r++) rows[r] = h; o.rows(rows); } };

  /* ---------- Engagement index and sample data ---------- */
  function index(set, wp = {}) {
    return [{step:'Write the engagement index', sheet:'Index', clear:true, tab:TITLE, ops:o => {
      o.cols({A:2, B:22, C:30, D:30, E:14, F:12, G:12});
      o.v('B2', 'Pipe audit file'); o.f('B2', {b:true, sz:15, c:TITLE});
      o.v('B4', [['Client name'],['Year end'],['Framework'],['Prepared by'],['Reviewer'],['Materiality basis'],['Senior / generic users']]);
      o.f('B4:B10', {b:true});
      const [y, m, d] = set.ye.split('-').map(Number);
      o.v('D4', [[set.client],[D.serial(y, m, d)],[set.framework],['PIPE'],[set.reviewer],[BASIS[set.entity]],[set.users]]);
      o.f('D4:D10', {bd:'grid', h:'Left'}); o.f('D4:D6', {fill:INPUT}); o.f('D8:D10', {fill:INPUT}); o.f('D5', {nf:DATE});
      heading(o, 'B12', 'Working papers');
      thead(o, 'B13', ['Ref','Working paper','Process','Status','Prepared','Reviewed']);
      o.v('B14', WPS.map(([ref, sheet, proc]) => { const w = wp[ref] || {}; return [ref, `=HYPERLINK("#${q(sheet)}!A1","${sheet}")`, proc, w.status || 'Not started', w.prep || '', w.rev || '']; }));
      o.f('B14:G23', {bd:'grid'}); o.f('F14:G23', {nf:DATE});
      o.cf('E14:E23', {text:'Reviewed', color:GOOD});
      o.cf('E14:E23', {text:'Prepared', color:WARN});
      note(o, 'B25', 'Edit the shaded cells to change the header on every working paper.');
    }}];
  }
  function indexRow(ref, status, which) {
    const row = 14 + WPS.findIndex(w => w[0] === ref);
    return o => { o.v(`E${row}`, status); o.v(`${which === 'rev' ? 'G' : 'F'}${row}`, today()); if (which !== 'rev') o.v(`G${row}`, ''); };
  }

  function sampleData() {
    return [
      {step:'Create the trial balance sheet', sheet:'TB', clear:true, tab:'#8A99A6', ops:o => {
        o.cols({A:10, B:30, C:18, D:30, E:12});
        o.v('A1', 'Trial balance · Harrow Precision Components (Pty) Ltd · 30 Sep 2026'); o.f('A1', {b:true});
        thead(o, 'A3', ['Account','Name','Balance','FS line','Class']);
        o.v('A4', D.TB);
        const last = 3 + D.TB.length;
        o.v(`B${last + 1}`, 'Total (should be nil)'); o.v(`C${last + 1}`, `=SUM(C4:C${last})`);
        o.f(`A${last + 1}:E${last + 1}`, {b:true, bd:'top'}); o.f(`C4:C${last + 1}`, {nf:ACC});
        note(o, 'A2', 'Demo data. Debits positive, credits negative.');
      }},
      {step:'Create the general ledger sheet', sheet:'GL', clear:true, tab:'#8A99A6', ops:o => {
        const lines = D.ledger();
        o.cols({A:11, B:11, C:9, D:26, E:36, F:15, G:15, H:9, I:10});
        o.v('A1', 'General ledger extract · FY2026'); o.f('A1', {b:true});
        note(o, 'A2', 'Demo extract of system sales, purchases and every manual journal. The full ledger has 48,316 lines.');
        thead(o, 'A3', ['Date','Journal','Account','Account name','Description','Debit','Credit','Source','User']);
        o.v('A4', lines);
        o.f(`A4:A${3 + lines.length}`, {nf:DATE}); o.f(`F4:G${3 + lines.length}`, {nf:ACC});
      }}
    ];
  }

  /* ---------- A1 Materiality ---------- */
  function materiality(C) {
    const S = 'A1 Materiality', tb = C.tb ? q(C.tbName) : null;
    const basisRow = {earnings:9, revenue:8, assets:11}[C.basis] || 9;
    const rate = {9:0.05, 8:0.01, 11:0.02}[basisRow];
    const B = C.bench || {};
    const val = (f, k) => tb ? f : B[k];
    const bases = [
      ['Revenue', 0.005, 0.01, val(`=-SUMIFS(${tb}!$C:$C,${tb}!$E:$E,"Revenue")`, 'revenue')],
      ['Net profit before tax', 0.05, 0.1, val(`=-(SUMIFS(${tb}!$C:$C,${tb}!$E:$E,"Revenue")+SUMIFS(${tb}!$C:$C,${tb}!$E:$E,"Expense"))`, 'pbt')],
      ["Shareholder's equity", 0.02, 0.05, val(`=-SUMIFS(${tb}!$C:$C,${tb}!$E:$E,"Equity")`, 'equity')],
      ['Total Assets', 0.01, 0.02, val(`=SUMIFS(${tb}!$C:$C,${tb}!$E:$E,"Asset")`, 'assets')],
      ['Total Liabilities', 0.01, 0.02, val(`=-SUMIFS(${tb}!$C:$C,${tb}!$E:$E,"Liability")`, 'liabilities')],
      ['Other', 0.005, 0.1, '']
    ];
    const why = {
      9:'Net profit before tax has been selected as the basis for materiality because Harrow is a profit-oriented, owner-managed manufacturer and its lenders focus on earnings. Profit is stable year on year (FY2025: R 13.9m), so no normalisation is needed. Revenue was considered and rejected: 1% of revenue is too high for an earnings-focused user group.',
      8:'Revenue has been selected as the basis for materiality because the users of the financial statements focus on revenue and activity levels. Profit was considered and rejected because it is not the primary measure for these users.',
      11:'Total assets has been selected as the basis for materiality because the users focus on the net asset position. Profit and revenue were considered and rejected as less relevant to these users.'
    }[basisRow];
    return [
      {step:'Pull benchmarks from the trial balance', sheet:S, clear:true, tab:TITLE, ops:o => {
        o.cols({A:2, B:3, C:28, D:12, E:8, F:8, G:9, H:12, I:16, J:16, K:16, L:16});
        header(o, 'A1', 'Materiality', 'L');
        o.v('B7', 'Basis for materiality'); o.v('G7', [['Min %','Max %','=D3','Basis used','Range low','Range high']]);
        o.f('B7:L7', {b:true, fill:HEAD, bd:'bottom'}); o.f('G7:L7', {h:'Right'}); o.f('I7', {nf:DATE}); o.f('J7', {h:'Center'});
        o.v('C8', bases.map(b => [b[0]])); o.v('G8', bases.map(b => [b[1], b[2], b[3]]));
        o.v('K8', bases.map((b, i) => [`=IF(I${8 + i}="","",I${8 + i}*G${8 + i})`, `=IF(I${8 + i}="","",I${8 + i}*H${8 + i})`]));
        o.f('G8:H13', {nf:PCT}); o.f('I8:I13', {nf:ACC}); o.f('K8:L13', {nf:ACC}); o.f('B13:L13', {bd:'bottom'});
        if (!tb) { o.f('I8:I12', {fill:INPUT}); note(o, 'C14', 'No trial balance sheet: benchmarks provided by the auditor.'); }
      }},
      {step:'Choose the basis and set materiality', sheet:S, ops:o => {
        o.v(`J${basisRow}`, '✓'); o.f(`J${basisRow}`, {b:true, h:'Center', c:GOOD});
        o.v('B17', 'Materiality will be set at'); o.v('H17', 'Rate'); o.v('I17', rate); o.v('J17', `=ROUNDDOWN(I${basisRow}*I17,-4)`);
        note(o, 'C18', `${BASIS[C.basis] || 'Profit before tax'} × rate, rounded down to the nearest R 10,000`);
        o.v('B19', '="Performance materiality will be set at "&TEXT(I19,"0%")&" of materiality"'); o.v('I19', 0.75); o.v('J19', '=J17*I19');
        o.v('B21', '="Clearly trivial is set at "&TEXT(I21,"0%")&" of materiality"'); o.v('I21', 0.05); o.v('J21', '=J17*I21');
        o.f('I17', {nf:PCT, fill:INPUT, bd:'box'}); o.f('I19', {nf:PCT0, fill:INPUT, bd:'box'}); o.f('I21', {nf:PCT0, fill:INPUT, bd:'box'});
        o.f('J17', {nf:ACC, b:true, bd:'box'}); o.f('J19', {nf:ACC, b:true}); o.f('J21', {nf:ACC, b:true});
      }},
      {step:'Write the rationale', sheet:S, select:'J17', ops:o => {
        para(o, 'B15', 'B15:L16', why, 24);
        para(o, 'B23', 'B23:L24', 'Performance materiality is set at the top of the firm range (75%) because the FY2025 file shows no significant misstatements and the finance team is unchanged. Clearly trivial follows the firm policy of 5% of materiality.', 18);
        note(o, 'B26', 'Sources: Materiality Policy 2026 §2.3 and §3.1 · Harrow FY2025 working papers A2');
        note(o, 'B27', 'Prepared by Pipe. Shaded cells are judgements you can change; the figures recalculate.');
      }}
    ];
  }

  /* ---------- A3 Journal testing ---------- */
  function journals(C) {
    const S = 'A3 Journals', gl = q(C.glName), js = D.journals(), last = 22 + js.length;
    return [
      {step:'Count the population', sheet:S, clear:true, tab:TITLE, ops:o => {
        o.cols({A:2, B:11, C:11, D:10, E:30, F:32, G:14, H:9, I:9, J:9, K:9, L:9, M:9, N:8, O:12});
        header(o, 'A3', 'Journal entry testing', 'O');
        heading(o, 'B7', 'Population');
        o.v('C8', [['GL lines in workbook extract'],['Manual journal lines'],['Manual journals analysed']]);
        o.v('I8', [[`=COUNTA(${gl}!$B$4:$B$10000)`],[`=COUNTIF(${gl}!$H$4:$H$10000,"Manual")`],[`=COUNTA(B23:B${last})`]]);
        o.f('I8:I10', {b:true, h:'Right'});
      }},
      {step:'Apply the risk criteria', sheet:S, ops:o => {
        heading(o, 'B12', 'Criteria (1 point each)');
        o.v('C13', [['Weekend posting'],['Round amount'],['Near year end'],['Post-close'],['Senior or generic user'],['Unusual pairing']]);
        o.v('F13', [['Posted on a Saturday or Sunday'],['A multiple of R 10,000'],['Within 3 days before year end'],['Posted after year end'],['User listed in I19'],["Pipe's judgement on the account combination"]]);
        o.v('C19', 'Senior or generic users'); o.v('I19', '=Index!$D$10'); o.f('I19', {fill:INPUT, b:true});
        note(o, 'F20', 'Source: Firm methodology §6.3 · ISA 240 ¶32');
      }},
      {step:'Score every manual journal', sheet:S, ops:o => {
        heading(o, 'B21', 'Manual journals');
        thead(o, 'B22', ['Journal','Date','User','Description','Account pairing','Amount','Weekend','Round','Near YE','Post-close','User','Pairing','Score','Result']);
        o.v('B23', js.map((j, i) => {
          const r = 23 + i;
          return [j.je, j.date, j.user, j.desc, j.pair, j.amt,
            `=IF(WEEKDAY(C${r},2)>5,1,0)`, `=IF(MOD(G${r},10000)=0,1,0)`, `=IF(AND(C${r}<=$D$3,C${r}>=$D$3-3),1,0)`, `=IF(C${r}>$D$3,1,0)`,
            `=IF(ISNUMBER(SEARCH(D${r},$I$19)),1,0)`, j.judg, `=SUM(H${r}:M${r})`, `=IF(N${r}>=3,"Investigate",IF(N${r}>=1,"Review","-"))`];
        }));
        o.f(`C23:C${last}`, {nf:DATE}); o.f(`G23:G${last}`, {nf:ACC}); o.f(`H23:N${last}`, {h:'Center'}); o.f(`M23:M${last}`, {fill:INPUT});
        o.f(`B23:O${last}`, {bd:'grid'}); o.f(`N23:N${last}`, {b:true});
        o.cf(`O23:O${last}`, {text:'Investigate', color:BAD, fill:BAD_FILL});
        o.cf(`O23:O${last}`, {text:'Review', color:WARN, fill:WARN_FILL});
      }},
      {step:'Summarise and conclude', sheet:S, select:`O23`, ops:o => {
        const s = last + 2;
        heading(o, `B${s}`, 'Summary');
        o.v(`C${s + 1}`, [['Journals analysed'],['Met one or more criteria'],['Met three or more (investigate)']]);
        o.v(`I${s + 1}`, [[`=COUNTA(B23:B${last})`],[`=COUNTIF(N23:N${last},">=1")`],[`=COUNTIF(N23:N${last},">=3")`]]);
        o.f(`I${s + 1}:I${s + 3}`, {b:true, h:'Right'});
        para(o, `B${s + 5}`, `B${s + 5}:O${s + 6}`, `="Of "&I${s + 1}&" manual journals analysed, "&I${s + 3}&" met three or more criteria and go to targeted testing in J1.1. JE-24117 and JE-24203 move R 450,000 of revenue across the year end, which links to the revenue cut-off risk in A2."`, 18);
        note(o, `B${s + 8}`, 'Prepared by Pipe. The pairing column is a judgement; every other criterion is a formula on the journal data.');
      }}
    ];
  }

  /* ---------- A2 Risk assessment ---------- */
  function risk(C) {
    const S = 'A2 Risk assessment';
    return [
      {step:'Analytical review against prior year', sheet:S, clear:true, tab:TITLE, ops:o => {
        o.cols({A:2, B:16, C:20, D:12, E:12, F:13, G:52, H:40, I:18, J:12});
        header(o, 'A2', 'Risk assessment');
        heading(o, 'B7', 'Monthly revenue');
        thead(o, 'B8', ['Month','FY2026','FY2025','Change','Change %']);
        o.v('B9', D.MONTHS.map((m, i) => [m, D.REV_CY[i], D.REV_PY[i], `=C${9 + i}-D${9 + i}`, `=IF(D${9 + i}=0,"",E${9 + i}/D${9 + i})`]));
        o.v('B21', [['Total', '=SUM(C9:C20)', '=SUM(D9:D20)', '=C21-D21', '=E21/D21']]);
        o.v('B22', 'September vs average of other months'); o.v('F22', '=C20/AVERAGE(C9:C19)-1');
        o.f('C9:E21', {nf:ACC}); o.f('F9:F22', {nf:PCT}); o.f('B21:F21', {b:true, bd:'top'}); o.f('B20:F20', {fill:WARN_FILL}); o.f('F22', {b:true, c:BAD});
      }},
      {step:'Link materiality', sheet:S, ops:o => {
        o.v('B24', [['Overall materiality'],['Performance materiality']]);
        if (C.a1) o.v('F24', [["='A1 Materiality'!$J$17"],["='A1 Materiality'!$J$19"]]);
        else { o.v('F24', [[C.om],[C.pm]]); o.f('F24:F25', {fill:INPUT}); note(o, 'G24', 'Provided by the auditor (no A1 in this workbook)'); }
        o.f('F24:F25', {nf:ACC, b:true});
      }},
      {step:'Rate each area and assertion', sheet:S, ops:o => {
        heading(o, 'B27', 'Risk register');
        thead(o, 'B28', ['Area','Assertion','Likelihood','Magnitude','Risk','Rationale','Planned response','Source']);
        o.v('B29', D.RISKS.map((r, i) => {
          const n = 29 + i;
          return [r.area, r.as, r.l, r.m,
            `=IF(D${n}="Presumed","Significant",IF(AND(D${n}="High",E${n}="High"),"Significant",IF(OR(D${n}="High",E${n}="High"),"High",IF(AND(D${n}="Low",E${n}="Low"),"Low","Medium"))))`,
            r.area === 'Journals' && !C.a3 ? 'Presumed significant under ISA 240. Journal testing (A3) has not been run in this workbook yet.' : D.RISK_WHY[r.area], r.resp, r.src];
        }));
        o.f('B29:I36', {bd:'grid', va:'Top'}); o.f('G29:H36', {w:true}); o.f('D29:E36', {fill:INPUT}); o.f('F29:F36', {b:true});
        o.dv('D29:D36', ['Presumed','High','Medium','Low']); o.dv('E29:E36', ['High','Medium','Low']);
        o.cf('F29:F36', {text:'Significant', color:BAD, fill:BAD_FILL});
        o.cf('F29:F36', {text:'High', color:WARN, fill:WARN_FILL});
        o.rows(Object.fromEntries(Array.from({length:8}, (_, i) => [29 + i, 46])));
      }},
      {step:'Summarise the assessment', sheet:S, select:'F29', ops:o => {
        o.v('B38', '="Risk register: "&COUNTIF(F29:F36,"Significant")&" significant, "&COUNTIF(F29:F36,"High")&" high, "&COUNTIF(F29:F36,"Medium")&" medium and "&COUNTIF(F29:F36,"Low")&" low. Significant and high risks are carried into B1 Sampling."');
        o.f('B38', {b:true});
        note(o, 'B40', 'Sources: ISA 315 (Revised) ¶31 · Firm risk library (manufacturing) · Harrow FY2025 file');
        note(o, 'B41', 'Prepared by Pipe. Likelihood and magnitude are judgements; the rating is a formula.');
      }}
    ];
  }

  /* ---------- B1 Sampling ---------- */
  function sampling(C) {
    const S = 'B1 Sampling', tb = C.tb ? q(C.tbName) : null, list = C.samples, start = 44, end = start + list.length - 1;
    const pmF = C.a1 ? "='A1 Materiality'!$J$19" : C.pm;
    return [
      {step:'Map GL accounts to audit areas', sheet:S, clear:true, tab:TITLE, ops:o => {
        o.cols({A:2, B:14, C:26, D:16, E:28, F:14, G:12, H:14, I:16, J:12, K:12, L:14});
        header(o, 'B1', 'Account mapping & sampling', 'L');
        heading(o, 'B7', 'Account mapping');
        thead(o, 'B8', ['Account','Name','TB balance','FS line','Audit area','Note']);
        o.v('B9', D.MAPPING.map(([acct, area, n], i) => {
          const r = 9 + i, bal = D.TB.find(t => t[0] === acct);
          return [acct, D.ACCT[acct], tb ? `=SUMIFS(${tb}!$C:$C,${tb}!$A:$A,B${r})` : bal[2], tb ? `=IFERROR(INDEX(${tb}!$D:$D,MATCH(B${r},${tb}!$A:$A,0)),"")` : bal[3], area, n];
        }));
        o.f('D9:D21', {nf:ACC}); o.f('F9:F21', {fill:INPUT}); o.f('B9:G21', {bd:'grid'});
        o.dv('F9:F21', ['Revenue','Inventory','Payables','PPE','Not sampled']);
        if (!tb) note(o, 'C22', 'No trial balance sheet: balances entered as values.');
      }},
      {step:'Set parameters and sample sizes', sheet:S, ops:o => {
        heading(o, 'B23', 'Parameters');
        o.v('C24', [['Performance materiality'],['Selection seed']]); o.v('I24', [[pmF],[20261102]]);
        o.f('I24', {nf:ACC, b:true}); if (!C.a1) o.f('I24', {fill:INPUT});
        heading(o, 'B27', 'Sampling plan');
        thead(o, 'B28', ['Area','Risk','Method','Population','Confidence','Sample size','Interval','Sample value','Coverage','Firm minimum','Check']);
        o.v('B29', D.AREAS.map((a, i) => {
          const r = 29 + i;
          const pop = {REV:'=-SUMIFS($D$9:$D$21,$F$9:$F$21,"Revenue")', INV:'=SUMIFS($D$9:$D$21,$F$9:$F$21,"Inventory")', PAY:'=-SUMIFS($D$9:$D$21,$F$9:$F$21,"Payables")', PPE:'=SUMIFS($D$9:$D$21,$F$9:$F$21,"PPE")', JE:C.a3 ? "=COUNTIF('A3 Journals'!$N$23:$N$52,\">=1\")" : 6}[a.id];
          return [a.name, C.a2 ? `='A2 Risk assessment'!$F$${a.a2Row}` : C.risks[a.id], a.method, pop, `=IF(C${r}="Significant",0.95,0.9)`, C.sizes[a.id],
            `=IF(D${r}="Monetary unit sampling",E${r}/G${r},"")`, `=SUMIF($C$${start}:$C$300,B${r},$G$${start}:$G$300)`, `=IF(OR(B${r}="Journals",E${r}=0),"",I${r}/E${r})`,
            `=IF(D${r}="Targeted (risk criteria)","n/a",VLOOKUP(C${r},$C$37:$D$40,2,FALSE))`, `=IF(K${r}="n/a","Targeted",IF(G${r}>=K${r},"OK","Below minimum"))`];
        }));
        o.f('E29:E32', {nf:ACC}); o.f('E33', {nf:'0 "journals"'}); o.f('F29:F33', {nf:PCT0}); o.f('H29:I33', {nf:ACC}); o.f('J29:J33', {nf:PCT});
        o.f('G29:G33', {fill:INPUT, b:true}); o.f('B29:L33', {bd:'grid'}); o.f('F29:L33', {h:'Right'});
        o.cf('L29:L33', {text:'Below', color:BAD, fill:BAD_FILL});
        heading(o, 'B36', 'Firm minimum sample sizes');
        o.v('C37', [['Significant', 25],['High', 15],['Medium', 8],['Low', 5]]); o.f('C37:D40', {bd:'grid'});
        note(o, 'F37', 'Source: Sampling Guidance §3.1 and §4.2');
        if (!C.a2) note(o, 'F38', 'Risk levels provided by the auditor (no A2 in this workbook).');
      }},
      {step:`Select ${list.length} items`, sheet:S, ops:o => {
        heading(o, 'B42', 'Selected items');
        thead(o, 'B43', ['Sample','Area','GL reference','Date','Description','Amount','Why selected']);
        o.v(`B${start}`, list.map(s => [s.id, D.AREAS.find(a => a.id === s.area).name, s.ref, s.date, s.party, s.amt, s.why]));
        o.f(`E${start}:E${end}`, {nf:DATE}); o.f(`G${start}:G${end + 1}`, {nf:ACC}); o.f(`B${start}:H${end}`, {bd:'grid'});
        o.v(`F${end + 1}`, 'Total'); o.v(`G${end + 1}`, `=SUM(G${start}:G${end})`); o.f(`F${end + 1}:G${end + 1}`, {b:true});
      }},
      {step:'Write the sampling memo', sheet:S, select:'G29', ops:o => {
        para(o, `B${end + 3}`, `B${end + 3}:L${end + 4}`, 'Revenue uses monetary unit sampling because the risk is overstatement and the population is large. Inventory takes every SKU above performance materiality, then a random draw from two value strata. Payables are tested for completeness through payments after year end. Journals are a targeted selection, starting with those scoring three or more in A3.', 30);
        note(o, `B${end + 6}`, 'Prepared by Pipe. Sample sizes are judgements; change them and re-run sampling to redraw the items.');
      }}
    ];
  }

  /* ---------- Detail testing (one sheet per area) ---------- */
  function detail(C, a) {
    const S = a.sheet, list = C.samples.filter(s => s.area === a.id), first = 24, last = first + list.length - 1, tot = last + 1;
    const b1 = a.b1Row;
    const ev = D.evalArea(a, list), pm = C.pm;
    const concl = {
      REV: ev.projected > pm ? `Revenue cut-off is not operating effectively. Known misstatement of ${R(ev.known)} and projected misstatement of ${R(ev.projected)} exceed performance materiality. Propose an adjustment for the cut-off error, extend cut-off testing to all invoices from 15 Sep to 15 Oct 2026, and report to those charged with governance.` : `No material misstatement identified in revenue. Known misstatement of ${R(ev.known)} goes to S1.`,
      INV: `Projected misstatement of ${R(ev.projected)} is below performance materiality. Record the count difference on S1. Inventory existence is fairly stated, subject to the NRV review.`,
      PAY: ev.known ? `One unrecorded liability of ${R(ev.known)} found. Below performance materiality, but extend the search for unrecorded liabilities to 30 Nov 2026 and record the item on S1.` : 'No unrecorded liabilities identified.',
      PPE: ev.missing ? 'Testing is incomplete: support for one addition is outstanding and has been requested from the client. Conclusion pending.' : 'Additions are supported and correctly classified.',
      JE: 'JE-24117 has no evidence of approval and links to the revenue cut-off exception in D3.1. Escalate to the engagement partner as a control deficiency over manual journals. The other journals had a valid rationale and support.'
    }[a.id];
    return [
      {step:`${a.ref} ${a.name}: set up the paper`, sheet:S, clear:true, tab:'#297AA0', ops:o => {
        o.cols({A:2, B:9, C:13, D:11, E:30, F:14, G:13, H:11, I:14, J:12, K:18, L:14, M:46});
        header(o, a.ref, a.full + ': tests of detail', 'M');
        heading(o, 'B7', 'Objective');
        para(o, 'B8', 'B8:M8', `To obtain sufficient appropriate evidence that ${a.name.toLowerCase()} is free from material misstatement for ${a.assert.toLowerCase()}, in response to the risk assessed in A2.`, 30);
        heading(o, 'B10', 'Population and selection');
        o.v('C11', [['Population'],['Sample size'],['Method'],['Performance materiality'],['Clearly trivial']]);
        o.v('I11', [
          [C.b1 ? `='B1 Sampling'!$E$${b1}` : (a.id === 'JE' ? 6 : a.popV)],
          [C.b1 ? `='B1 Sampling'!$G$${b1}` : list.length],
          [C.b1 ? `='B1 Sampling'!$D$${b1}` : a.method],
          [C.a1 ? "='A1 Materiality'!$J$19" : C.b1 ? "='B1 Sampling'!$I$24" : C.pm],
          [C.a1 ? "='A1 Materiality'!$J$21" : '=I14/0.75*0.05']
        ]);
        o.f('I11', {nf:a.id === 'JE' ? '0 "journals"' : ACC}); o.f('I14:I15', {nf:ACC}); o.f('I11:I15', {h:'Right'});
        if (!C.b1) note(o, 'C16', 'No B1 in this workbook: default selection drawn by Pipe.');
        heading(o, 'B17', 'Procedure');
        o.v('B18', [[`1. Obtained ${a.doc.toLowerCase()} for each sample from the client portal.`],['2. Pipe extracted the document fields and matched them to the GL entry.'],[`3. Checked that ${a.checks}.`],['4. Investigated differences and evaluated them against materiality.']]);
      }},
      {step:`${a.ref} ${a.name}: match ${list.length} samples to documents`, sheet:S, ops:o => {
        heading(o, 'B22', 'Results');
        thead(o, 'B23', ['Sample','GL reference','Date','Description','GL amount','Document','Doc date','Doc amount','Difference','Result','Misstatement','Comment']);
        o.v(`B${first}`, list.map((s, i) => {
          const r = first + i, t = D.test(s);
          const mis = {none:0, zero:0, full:`=F${r}`, diff:`=J${r}`}[t.mis];
          return [s.id, s.ref, s.date, s.party, s.amt, t.doc, t.docDate, t.docAmt, `=IF(I${r}="","",F${r}-I${r})`, t.result, mis, t.note];
        }));
        o.f(`D${first}:D${last}`, {nf:DATE}); o.f(`H${first}:H${last}`, {nf:DATE}); o.f(`F${first}:F${tot}`, {nf:ACC}); o.f(`I${first}:J${last}`, {nf:ACC}); o.f(`L${first}:L${tot}`, {nf:ACC});
        o.f(`B${first}:M${last}`, {bd:'grid'}); o.f(`M${first}:M${last}`, {w:true}); o.f(`K${first}:K${last}`, {fill:INPUT});
        o.dv(`K${first}:K${last}`, ['Agreed','Exception','No support']);
        o.cf(`K${first}:K${last}`, {text:'Exception', color:BAD, fill:BAD_FILL});
        o.cf(`K${first}:K${last}`, {text:'No support', color:WARN, fill:WARN_FILL});
        o.v(`E${tot}`, 'Total'); o.v(`F${tot}`, `=SUM(F${first}:F${last})`); o.v(`L${tot}`, `=SUM(L${first}:L${last})`); o.f(`B${tot}:M${tot}`, {b:true, bd:'top'});
      }},
      {step:`${a.ref} ${a.name}: evaluate and conclude`, sheet:S, select:'L16', ops:o => {
        o.v('K10', 'Results summary'); o.f('K10', {b:true, c:TITLE});
        o.v('K11', [['Agreed', `=COUNTIF(K${first}:K${last},"Agreed")`], ['Exceptions', `=COUNTIF(K${first}:K${last},"Exception")`], ['No support', `=COUNTIF(K${first}:K${last},"No support")`],
          ['Known misstatement', `=L${tot}`], ['Projected', a.project ? `=IF(F${tot}=0,0,L14*I11/F${tot})` : '=L14'], ['Result', '=IF(L15>I14,"Exceeds PM",IF(L14>I15,"Above trivial","Clearly trivial"))']]);
        o.f('K11:L16', {bd:'grid'}); o.f('L14:L15', {nf:ACC}); o.f('L11:L16', {h:'Right', b:true});
        o.cf('L16', {text:'Exceeds', color:BAD, fill:BAD_FILL});
        heading(o, `B${tot + 2}`, 'Conclusion');
        para(o, `B${tot + 3}`, `B${tot + 3}:M${tot + 4}`, concl, 22);
        note(o, `B${tot + 6}`, `Projection: ${a.project ? 'known misstatement × population ÷ sample value (ratio method).' : 'not projected; known misstatement only.'} Sources: Sampling Guidance §7 · ToD WP template`);
        note(o, `B${tot + 7}`, 'Prepared by Pipe. Results are judgements you can change; the summary recalculates.');
      }}
    ];
  }
  const detailAll = C => D.AREAS.flatMap(a => detail(C, a));

  /* ---------- S1 Summary of misstatements ---------- */
  function summary(C) {
    const S = 'S1 Misstatements';
    return [
      {step:'Collect misstatements from each paper', sheet:S, clear:true, tab:TITLE, ops:o => {
        o.cols({A:2, B:10, C:24, D:13, E:16, F:16, G:16, H:12, I:14, J:12});
        header(o, 'S1', 'Summary of misstatements');
        thead(o, 'B7', ['WP','Area','Exceptions','Known','Projected','Status']);
        o.v('B8', D.AREAS.map((a, i) => {
          const r = 8 + i, s = q(a.sheet), m = (C.manual || {})[a.id] || {};
          return [a.ref, a.name,
            C.det[a.id] ? `=${s}!$L$12+${s}!$L$13` : (m.exc || 0),
            C.det[a.id] ? `=${s}!$L$14` : (m.known || 0),
            C.det[a.id] ? `=${s}!$L$15` : (m.projected || 0),
            `=IF(F${r}>$F$16,"Above PM",IF(E${r}>$F$17,"Above trivial","Trivial"))`];
        }));
        o.v('B13', [['', 'Total', '=SUM(D8:D12)', '=SUM(E8:E12)', '=SUM(F8:F12)', '']]);
        o.f('E8:F13', {nf:ACC}); o.f('B8:G12', {bd:'grid'}); o.f('B13:G13', {b:true, bd:'top'});
        if (D.AREAS.some(a => !C.det[a.id])) { o.f('D8:F12', {fill:INPUT}); note(o, 'C14', 'Some figures provided by the auditor (detail testing sheets not in this workbook).'); }
        o.cf('G8:G12', {text:'Above PM', color:BAD, fill:BAD_FILL});
      }},
      {step:'Compare with materiality', sheet:S, select:'F13', ops:o => {
        o.v('C15', [['Overall materiality'],['Performance materiality'],['Clearly trivial']]);
        o.v('F15', C.a1 ? [["='A1 Materiality'!$J$17"],["='A1 Materiality'!$J$19"],["='A1 Materiality'!$J$21"]] : [[C.om],[C.pm],[C.ct]]);
        o.v('C19', 'Headroom to overall materiality'); o.v('F19', '=F15-F13');
        o.f('F15:F19', {nf:ACC}); o.f('F19', {b:true}); if (!C.a1) o.f('F15:F17', {fill:INPUT});
        o.v('B21', '=IF(F13>F15,"Projected misstatement exceeds overall materiality. Adjustments are needed before the opinion is signed.",IF(F13>F16,"Projected misstatement exceeds performance materiality. Consider further procedures.","Projected misstatement is below performance materiality."))');
        o.f('B21', {b:true});
        para(o, 'B23', 'B23:J24', 'Most of the projected misstatement comes from the revenue cut-off error in D3.1 (REV-07, R 612,500). Propose that adjustment to management first; the remaining items are below performance materiality individually.', 18);
        note(o, 'B26', 'Prepared by Pipe. Figures link to each detail testing paper and recalculate.');
      }}
    ];
  }

  function signoff(initials) { return o => { o.v('I3', initials); o.v('J3', today()); o.f('I3:J3', {b:true, c:GOOD}); }; }

  g.PipeWP = {ORDER, WPS, BASIS, posFor, index, indexRow, sampleData, materiality, journals, risk, sampling, detail, detailAll, summary, signoff, today};
})(window);
