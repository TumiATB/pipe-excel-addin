/* Workbook writer. Builders produce plain ops; XL applies them to Excel, or to an
   in-browser preview workbook when the pane is opened outside Office. */
(function (g) {
  const colNum = c => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  function parse(a) {
    const [p, q] = a.split(':'), m1 = /^([A-Z]+)(\d+)$/.exec(p), m2 = /^([A-Z]+)(\d+)$/.exec(q || p);
    return {c1:colNum(m1[1]), r1:+m1[2], c2:colNum(m2[1]), r2:+m2[2]};
  }
  const dims = a => { const x = parse(a); return [x.r2 - x.r1 + 1, x.c2 - x.c1 + 1]; };
  const fill = (a, v) => { const [r, c] = dims(a); return Array.from({length:r}, () => Array(c).fill(v)); };
  // A single-cell address with a bigger matrix grows to fit it
  function fit(a, m) {
    if (a.includes(':')) return a;
    const x = parse(a);
    return `${colName(x.c1)}${x.r1}:${colName(x.c1 + m[0].length - 1)}${x.r1 + m.length - 1}`;
  }

  // Ops builder handed to every working-paper section
  function ops() {
    const list = [];
    const mat = v => Array.isArray(v) ? (Array.isArray(v[0]) ? v : [v]) : [[v]];
    return {
      list,
      v(a, v) { const m = mat(v); list.push({a:fit(a, m), v:m}); },
      f(a, fmt) { list.push({a, fmt}); },
      dv(a, items) { list.push({a, dv:items}); },
      cf(a, rule) { list.push({a, cf:rule}); },
      cols(c) { list.push({cols:c}); },
      rows(r) { list.push({rows:r}); }
    };
  }

  /* ---------- Excel ---------- */
  const BORDER = '#9FB3C2';
  function edges(r, a, kind) {
    const [rows, cols] = dims(a);
    let names = kind === 'bottom' ? ['EdgeBottom'] : kind === 'top' ? ['EdgeTop'] : ['EdgeTop','EdgeBottom','EdgeLeft','EdgeRight'];
    if (kind === 'grid') { if (rows > 1) names.push('InsideHorizontal'); if (cols > 1) names.push('InsideVertical'); }
    names.forEach(n => { const b = r.format.borders.getItem(n); b.style = 'Continuous'; b.color = BORDER; if (kind === 'thick') b.weight = 'Medium'; });
  }
  function applyFmt(r, a, f) {
    if (f.m) r.merge(false);
    if (f.b != null) r.format.font.bold = f.b;
    if (f.i != null) r.format.font.italic = f.i;
    if (f.c) r.format.font.color = f.c;
    if (f.sz) r.format.font.size = f.sz;
    if (f.fill) r.format.fill.color = f.fill;
    if (f.nf) r.numberFormat = fill(a, f.nf);
    if (f.w != null) r.format.wrapText = f.w;
    if (f.h) r.format.horizontalAlignment = f.h;
    if (f.va) r.format.verticalAlignment = f.va;
    if (f.bd) edges(r, a, f.bd);
  }
  const ExcelXL = {
    mode: 'excel',
    async sheets() {
      return Excel.run(async ctx => { const ws = ctx.workbook.worksheets; ws.load('items/name'); await ctx.sync(); return ws.items.map(w => w.name); });
    },
    async exists(name) {
      return Excel.run(async ctx => { const ws = ctx.workbook.worksheets.getItemOrNullObject(name); await ctx.sync(); return !ws.isNullObject; });
    },
    async read(name, a) {
      return Excel.run(async ctx => { const r = ctx.workbook.worksheets.getItem(name).getRange(a); r.load('values'); await ctx.sync(); return r.values; });
    },
    async activeSheet() {
      return Excel.run(async ctx => { const ws = ctx.workbook.worksheets.getActiveWorksheet(); ws.load('name'); await ctx.sync(); return ws.name; });
    },
    async activate(name) {
      return Excel.run(async ctx => { const ws = ctx.workbook.worksheets.getItem(name); ws.activate(); ws.getRange('A1').select(); await ctx.sync(); });
    },
    // opt: {clear, tab, pos, select}
    async apply(name, list, opt = {}) {
      return Excel.run(async ctx => {
        let ws = ctx.workbook.worksheets.getItemOrNullObject(name);
        await ctx.sync();
        if (ws.isNullObject) {
          ws = ctx.workbook.worksheets.add(name);
          if (opt.pos != null) {
            const all = ctx.workbook.worksheets; all.load('items/name'); await ctx.sync();
            ws.position = Math.min(opt.pos(all.items.map(w => w.name)), all.items.length - 1);
          }
          ws.showGridlines = false;
        } else if (opt.clear) {
          const used = ws.getRange();
          used.unmerge(); used.clear(); used.conditionalFormats.clearAll(); used.dataValidation.clear();
          ws.showGridlines = false;
        }
        if (opt.tab) ws.tabColor = opt.tab;
        for (const op of list) {
          if (op.cols) { for (const [c, w] of Object.entries(op.cols)) ws.getRange(`${c}:${c}`).format.columnWidth = w * 6; continue; }
          if (op.rows) { for (const [r, h] of Object.entries(op.rows)) ws.getRange(`${r}:${r}`).format.rowHeight = h; continue; }
          const r = ws.getRange(op.a);
          if (op.v) r.formulas = op.v;
          if (op.fmt) applyFmt(r, op.a, op.fmt);
          if (op.dv) { r.dataValidation.clear(); r.dataValidation.rule = {list:{inCellDropDown:true, source:op.dv.join(',')}}; }
          if (op.cf) {
            const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
            cf.textComparison.format.font.color = op.cf.color;
            if (op.cf.fill) cf.textComparison.format.fill.color = op.cf.fill;
            cf.textComparison.rule = {operator:Excel.ConditionalTextOperator.contains, text:op.cf.text};
          }
        }
        if (!opt.quiet) { ws.activate(); if (opt.select) ws.getRange(opt.select).select(); }
        await ctx.sync();
      });
    }
  };

  /* ---------- Preview workbook (outside Excel) ---------- */
  const book = {order:[], sheets:{}, active:null};
  let onChange = () => {};
  const PreviewXL = {
    mode: 'preview',
    book,
    onChange(fn) { onChange = fn; },
    async sheets() { return [...book.order]; },
    async exists(name) { return !!book.sheets[name]; },
    async read(name, a) {
      const s = book.sheets[name], x = parse(a), out = [];
      for (let r = x.r1; r <= x.r2; r++) { const row = []; for (let c = x.c1; c <= x.c2; c++) { const cell = s && s.cells[colName(c) + r]; row.push(cell ? cell.v : ''); } out.push(row); }
      return out;
    },
    async activeSheet() { return book.active; },
    async activate(name) { book.active = name; onChange(); },
    async apply(name, list, opt = {}) {
      let s = book.sheets[name];
      if (!s) {
        s = book.sheets[name] = {cells:{}, cols:{}, rows:{}, tab:null};
        const pos = opt.pos ? opt.pos(book.order) : book.order.length;
        book.order.splice(Math.min(pos, book.order.length), 0, name);
      } else if (opt.clear) { s.cells = {}; s.cols = {}; s.rows = {}; }
      if (opt.tab) s.tab = opt.tab;
      for (const op of list) {
        if (op.cols) { Object.assign(s.cols, op.cols); continue; }
        if (op.rows) { Object.assign(s.rows, op.rows); continue; }
        const x = parse(op.a);
        for (let r = x.r1; r <= x.r2; r++) for (let c = x.c1; c <= x.c2; c++) {
          const k = colName(c) + r, cell = s.cells[k] || (s.cells[k] = {v:'', f:{}});
          if (op.v) cell.v = op.v[r - x.r1][c - x.c1];
          if (op.fmt) Object.assign(cell.f, op.fmt, {m:op.fmt.m ? op.a : cell.f.m});
          if (op.dv) cell.dv = op.dv;
        }
      }
      if (!opt.quiet) book.active = name;
      onChange();
    }
  };

  g.XLUtil = {ops, parse, colName, colNum, dims};
  g.makeXL = inExcel => inExcel ? ExcelXL : PreviewXL;
})(window);
