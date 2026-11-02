// Builds the full-audit workbook with the add-in's own builders and dumps it as JSON,
// so test/check_formulas.py can recalculate it in a real formula engine.
const fs = require('fs'), path = require('path'), vm = require('vm');
global.window = global;
for (const f of ['data.js', 'xl.js', 'wp.js']) vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'js', f), 'utf8'), {filename:f});
const D = PipeData, WP = PipeWP, XL = makeXL(false);

async function write(secs) {
  for (const s of secs) { const o = XLUtil.ops(); s.ops(o); await XL.apply(s.sheet, o.list, {clear:s.clear, pos:WP.posFor(s.sheet)}); }
}
(async () => {
  const sizes = Object.fromEntries(D.AREAS.map(a => [a.id, a.size]));
  const set = {client:'Harrow Precision Components (Pty) Ltd', ye:'2026-09-30', framework:'IFRS for SMEs', entity:'earnings', reviewer:'LD', users:'cfo.vdm, admin'};
  await write(WP.index(set));
  await write(WP.sampleData());
  const C = {tbName:'TB', glName:'GL', basis:'earnings', sizes, tb:true, gl:true, risks:{}, det:{}, om:740000, pm:555000, ct:37000, samples:D.samples(sizes)};
  await write(WP.materiality(C)); C.a1 = true;
  await write(WP.journals(C)); C.a3 = true;
  await write(WP.risk(C)); C.a2 = true;
  await write(WP.sampling(C)); C.b1 = true;
  await write(WP.detailAll(C)); D.AREAS.forEach(a => C.det[a.id] = true);
  await write(WP.summary(C));
  const out = {order:XL.book.order, sheets:{}};
  for (const n of XL.book.order) out.sheets[n] = Object.fromEntries(Object.entries(XL.book.sheets[n].cells).filter(([, c]) => c.v !== '' && c.v != null).map(([k, c]) => [k, c.v]));
  fs.writeFileSync(process.argv[2] || 'book.json', JSON.stringify(out));
  console.log('sheets', out.order.length);
})();
