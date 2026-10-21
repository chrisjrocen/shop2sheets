// Local test harness: runs Code.gs in Node with minimal Apps Script mocks.
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

class FakeSheet {
  constructor() { this.data = []; this.maxRows = 1000; }
  _cell(r, c) { return (this.data[r - 1] && this.data[r - 1][c - 1] !== undefined) ? this.data[r - 1][c - 1] : ''; }
  _set(r, c, v) { while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getLastRow() { let last = 0; this.data.forEach((row, i) => { if (row.some(v => v !== '' && v !== undefined)) last = i + 1; }); return last; }
  getLastColumn() { let last = 0; this.data.forEach(row => row.forEach((v, i) => { if (v !== '' && v !== undefined) last = Math.max(last, i + 1); })); return last; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(after, n) { if (after < this.data.length) this.data.splice(after, 0, ...Array.from({ length: n }, () => [])); this.maxRows += n; }
  insertRowsBefore(row, n) { this.insertRowsAfter(row - 1, n); }
  appendRow(vals) { this.data.splice(this.getLastRow(), 0, vals.slice()); }
  setFrozenRows() {} setColumnWidth() {} deleteRows() {}
  getConditionalFormatRules() { return this.rules || []; }
  setConditionalFormatRules(r) { this.rules = r; }
  insertColumnAfter(col) { this.data.forEach(row => { while (row.length < col) row.push(''); row.splice(col, 0, ''); }); }
  getRange(r, c, nr = 1, nc = 1) {
    const s = this;
    return {
      getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push(s._cell(r + i, c + j)); out.push(row); } return out; },
      setValues(vals) { vals.forEach((row, i) => row.forEach((v, j) => s._set(r + i, c + j, v))); return this; },
      setValue(v) { s._set(r, c, v); return this; },
      setNumberFormat() { return this; }, setFontWeight() { return this; }, setBackground() { return this; },
      setDataValidation(v) { s.validation = s.validation || {}; s.validation[r] = v; return this; },
      sort(specs) {
        const key = v => v instanceof Date ? v.getTime() : (v === '' || v === undefined ? null : v);
        const rows = this.getValues();
        rows.sort((a, b) => {
          for (const sp of specs) {
            const x = key(a[sp.column - c]), y = key(b[sp.column - c]);
            if (x === y) continue;
            if (x === null) return 1; if (y === null) return -1; // blanks last
            return (x < y ? -1 : 1) * (sp.ascending ? 1 : -1);
          }
          return 0;
        });
        this.setValues(rows); s.sorted = (s.sorted || 0) + 1; return this;
      }
    };
  }
}

const props = {};
const sheets = {};
const calls = [];
const ctx = {
  console, Date, JSON, Math, Object, Array, String, Number, Error, isNaN,
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] === undefined ? null : props[k], setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
  Utilities: { sleep() {} },
  SpreadsheetApp: {
    getActive: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet()) }),
    newConditionalFormatRule: () => { const rule = {}; const b = { whenFormulaSatisfied: f => { rule.f = f; return b; }, setBackground: () => b, setRanges: () => b, build: () => ({ getBooleanCondition: () => ({ getCriteriaValues: () => [rule.f] }) }) }; return b; },
    newDataValidation: () => { const v = {}; const b = { requireValueInList: l => { v.list = l; return b; }, setAllowInvalid: () => b, build: () => v }; return b; }
  },
  Session: { getEffectiveUser: () => ({ getEmail: () => ctx.__user || '' }) },
  ScriptApp: {
    triggers: [],
    getProjectTriggers() { return this.triggers; },
    deleteTrigger(t) { this.triggers = this.triggers.filter(x => x !== t); },
    newTrigger(handler) {
      const calls = [];
      const b = new Proxy({}, { get: (_, m) => m === 'create' ? () => { ctx.ScriptApp.triggers.push({ handler, calls, getHandlerFunction: () => handler }); } : (...a) => { calls.push([m, ...a]); return b; } });
      return b;
    }
  },
  UrlFetchApp: { fetch: (...a) => ctx.__fetch(...a) }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync('Code.gs', 'utf8'), ctx, { filename: 'Code.gs' });
const run = code => vm.runInContext(code, ctx);
const call = (fn, ...args) => { ctx.__args = args; return run(fn + '(...__args)'); };
// const declarations are not context properties, so read them via eval in context
// JSON round-trip so arrays belong to this realm (deepStrictEqual rejects cross-realm prototypes)
const J = x => JSON.parse(JSON.stringify(x));
const ORDER_HEADERS = J(run('ORDER_HEADERS')), DRAFT_HEADERS = J(run('DRAFT_HEADERS')), ABANDONED_HEADERS = J(run('ABANDONED_HEADERS'));
const ORDER_LAYOUT = J(run('ORDER_LAYOUT')), ABANDONED_LAYOUT = J(run('ABANDONED_LAYOUT'));

let passed = 0;
const test = (name, fn) => { try { fn(); passed++; console.log('ok   ' + name); } catch (e) { console.log('FAIL ' + name + '\n     ' + e.message); process.exitCode = 1; } };


test('normalizePhone_', () => {
  const n = p => call('normalizePhone_', p, '256');
  assert.strictEqual(n('0772 123 456'), '+256772123456');
  assert.strictEqual(n('256772123456'), '+256772123456');
  assert.strictEqual(n('00256772123456'), '+256772123456');
  assert.strictEqual(n('772123456'), '+256772123456');
  assert.strictEqual(n('+1 (415) 555-0100'), '+14155550100');
  assert.strictEqual(n(null), '');
});
test('runsOf_', () => {
  assert.deepStrictEqual(J(call('runsOf_', [5, 1, 2, 3, 7, 8])), [{ start: 1, width: 3 }, { start: 5, width: 1 }, { start: 7, width: 2 }]);
});
test('layouts: Remarks is owner column, never managed', () => {
  assert.ok(ORDER_LAYOUT.includes('Remarks'));
  [ORDER_HEADERS, DRAFT_HEADERS, ABANDONED_HEADERS].forEach(h => { assert.ok(!h.includes('Remarks')); assert.ok(h.includes('Shopify ID')); });
});

const mb = (amount, cur = 'UGX') => ({ presentmentMoney: { amount: String(amount), currencyCode: cur } });
const cfg = { country_code: '256' };
const order = {
  legacyResourceId: '5550001', name: '#1042', createdAt: '2026-10-14T03:00:00Z', cancelledAt: null,
  displayFinancialStatus: 'PENDING', displayFulfillmentStatus: 'UNFULFILLED', phone: null,
  customer: { displayName: 'Cust Name' },
  shippingAddress: { name: 'Jane Doe', phone: '0772123456', address1: 'Plot 1', address2: 'Ntinda', city: 'Kampala' },
  billingAddress: null, totalPriceSet: mb(105000),
  lineItems: { pageInfo: { hasNextPage: false }, nodes: [{ title: 'Shirt', quantity: 2, variantTitle: 'M' }, { title: 'Cap', quantity: 1, variantTitle: 'Default Title' }] }
};
test('helpers: locationOf_, clientName_, orderStatus_, itemsSummary_', () => {
  assert.strictEqual(call('locationOf_', { address1: 'Plot 1', address2: '', city: 'Kampala' }), 'Plot 1, Kampala');
  assert.strictEqual(call('locationOf_', null), '');
  assert.strictEqual(call('clientName_', {}, { name: 'B' }, { displayName: 'C' }), 'B');
  assert.strictEqual(call('clientName_', null, null, { displayName: 'C' }), 'C');
  assert.strictEqual(call('orderStatus_', order), 'Unpaid · Unfulfilled');
  assert.strictEqual(call('orderStatus_', Object.assign({}, order, { displayFinancialStatus: 'PAID', displayFulfillmentStatus: 'FULFILLED' })), 'Paid · Fulfilled');
  assert.strictEqual(call('orderStatus_', Object.assign({}, order, { cancelledAt: 'x' })), 'Cancelled');
  assert.strictEqual(call('itemsSummary_', { nodes: [{ name: 'Shirt - M', title: 'Shirt', quantity: 1, variantTitle: 'M' }] }).text, '1x Shirt - M');
  assert.strictEqual(call('itemsSummary_', { pageInfo: { hasNextPage: true }, nodes: [{ title: 'A', quantity: 1 }] }).text, '1x A; ...');
});
test('orderToRecord_', () => {
  const r = call('orderToRecord_', order, cfg);
  assert.strictEqual(r['Order ID'], '#1042');
  assert.strictEqual(r['Shopify ID'], '5550001');
  assert.strictEqual(r['Client name'], 'Jane Doe');
  assert.strictEqual(r['Items ordered'], '2x Shirt (M); 1x Cap');
  assert.strictEqual(r['Price'], 105000);
  assert.strictEqual(r['Phone'], '+256772123456');
  assert.strictEqual(r['Location'], 'Plot 1, Ntinda, Kampala');
  assert.strictEqual(r['Status'], 'Unpaid · Unfulfilled');
  assert.ok(r['Date placed'] instanceof Date);
  assert.deepStrictEqual(Object.keys(r).sort(), ORDER_HEADERS.slice().sort());
});
const draft = Object.assign({}, order, { legacyResourceId: '777', name: '#D12', status: 'OPEN', order: null });
test('draftToRecord_', () => {
  const r = call('draftToRecord_', draft, cfg);
  assert.strictEqual(r['Draft ID'], '#D12');
  assert.strictEqual(r['Status'], 'Open');
  assert.strictEqual(r['Order #'], '');
  assert.deepStrictEqual(Object.keys(r).sort(), DRAFT_HEADERS.slice().sort());
  const done = call('draftToRecord_', Object.assign({}, draft, { status: 'COMPLETED', order: { name: '#1050' } }), cfg);
  assert.strictEqual(done['Status'], 'Completed');
  assert.strictEqual(done['Order #'], '#1050');
});
const checkout = {
  id: 'gid://shopify/AbandonedCheckout/9001', name: '#9001', createdAt: '2026-10-14T01:00:00Z', completedAt: null,
  abandonedCheckoutUrl: 'https://shop.example/checkouts/cn/abcDEF123456xyz/recover?key=k',
  customer: { displayName: 'Sam', defaultEmailAddress: { emailAddress: 's@x.co' }, defaultPhoneNumber: { phoneNumber: '+256700111222' } },
  shippingAddress: null, billingAddress: null, totalPriceSet: mb(50000),
  lineItems: { pageInfo: { hasNextPage: false }, nodes: [{ title: 'Shoe', quantity: 1, variantTitle: '42' }] }
};
test('checkoutToRecord_', () => {
  const r = call('checkoutToRecord_', checkout, cfg);
  assert.strictEqual(r['Checkout ID'], '#9001');
  assert.strictEqual(r['Shopify ID'], '9001');
  assert.strictEqual(r['Status'], 'Abandoned');
  assert.strictEqual(r['Client name'], 'Sam');
  assert.strictEqual(r['Order #'], undefined);
  assert.strictEqual(r['Remarks'], undefined);
  const rec = call('checkoutToRecord_', Object.assign({}, checkout, { completedAt: '2026-10-14T02:00:00Z' }), cfg);
  assert.strictEqual(rec['Status'], 'Recovered');
  assert.ok(rec['Recovered at'] instanceof Date);
  assert.strictEqual(call('checkoutToRecord_', Object.assign({}, checkout, { customer: null }), cfg), null);
});

test('upsert_: append, update, idempotent, Remarks survive sort', () => {
  const sh = new FakeSheet();
  call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  assert.deepStrictEqual(sh.data[0], ORDER_LAYOUT);
  const remarksCol = ORDER_LAYOUT.indexOf('Remarks') + 1;
  const a = call('orderToRecord_', order, cfg);
  const b = call('orderToRecord_', Object.assign({}, order, { legacyResourceId: '5550002', name: '#1043' }), cfg);
  let r = call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [a, b]);
  assert.deepStrictEqual([r.added, r.updated], [2, 0]);
  sh._set(2, remarksCol, 'call first'); sh._set(3, remarksCol, 'gift');
  // owner sorts: swap rows
  const r2 = sh.data[1].slice(), r3 = sh.data[2].slice(); sh.data[1] = r3; sh.data[2] = r2;
  r = call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [a, b]);
  assert.deepStrictEqual([r.added, r.updated], [0, 0], 'idempotent');
  const paid = call('orderToRecord_', Object.assign({}, order, { displayFinancialStatus: 'PAID' }), cfg);
  r = call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [paid, paid]);
  assert.deepStrictEqual([r.added, r.updated], [0, 1]);
  const row = sh.data.findIndex(x => x[ORDER_LAYOUT.indexOf('Shopify ID')] === '5550001') + 1;
  assert.strictEqual(sh._cell(row, ORDER_LAYOUT.indexOf('Status') + 1), 'Paid · Unfulfilled');
  assert.strictEqual(sh._cell(row, remarksCol), 'call first');
  assert.strictEqual(sh.getLastRow(), 3, 'no duplicates');
});
test('ensureHeaders_/ensureOwnerColumns_: reorder and deleted Remarks', () => {
  const sh = new FakeSheet();
  call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  const h = sh.data[0]; [h[0], h[1]] = [h[1], h[0]];
  const m = call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  assert.strictEqual(m['Order ID'], 1); assert.strictEqual(m['Date placed'], 2);
  sh.data[0][ORDER_LAYOUT.indexOf('Remarks')] = '';
  call('ensureOwnerColumns_', sh, ORDER_LAYOUT);
  assert.ok(sh.data[0].includes('Remarks'));
  // records still land under the right headers
  call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [call('orderToRecord_', order, cfg)]);
  const m2 = call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  assert.strictEqual(sh._cell(2, m2['Order ID']), '#1042');
});
test('upsert_ abandoned: update never overwrites Order #', () => {
  const sh = new FakeSheet();
  const done = Object.assign({}, checkout, { completedAt: '2026-10-14T02:00:00Z' });
  call('upsert_', sh, ABANDONED_HEADERS, 'Shopify ID', [call('checkoutToRecord_', done, cfg)]);
  const m = call('ensureHeaders_', sh, ABANDONED_HEADERS);
  sh._set(2, m['Order #'], '#1050');
  const r = call('upsert_', sh, ABANDONED_HEADERS, 'Shopify ID', [call('checkoutToRecord_', Object.assign({}, done, { totalPriceSet: mb(60000) }), cfg)]);
  assert.strictEqual(r.updated, 1);
  assert.strictEqual(sh._cell(2, m['Order #']), '#1050');
  assert.strictEqual(sh._cell(2, m['Price']), 60000);
});

const resp = (code, body) => ({ getResponseCode: () => code, getContentText: () => typeof body === 'string' ? body : JSON.stringify(body) });
test('gql_: refreshes token on 401, backs off on THROTTLED', () => {
  Object.assign(props, { SHOPIFY_STORE: 'demo', CLIENT_ID: 'id', CLIENT_SECRET: 'sec' });
  let n = 0, tokens = 0;
  ctx.__fetch = (url, opts) => {
    if (url.endsWith('/admin/oauth/access_token')) { tokens++; assert.strictEqual(opts.payload.grant_type, 'client_credentials'); return resp(200, { access_token: 't' + tokens, expires_in: 86399 }); }
    n++;
    if (n === 1) return resp(401, 'unauthorized');
    if (n === 2) return resp(200, { errors: [{ extensions: { code: 'THROTTLED' }, message: 'slow' }] });
    return resp(200, { data: { ok: true } });
  };
  assert.deepStrictEqual(J(call('gql_', '{ shop { name } }', {})), { ok: true });
  assert.ok(tokens >= 2);
});
test('gql_: ACCESS_DENIED and auth failure', () => {
  props.TOKEN = 'x'; props.TOKEN_EXPIRES = String(Date.now() + 86400000);
  ctx.__fetch = () => resp(200, { errors: [{ extensions: { code: 'ACCESS_DENIED' }, message: 'not approved to access the Customer object' }] });
  assert.throws(() => call('gql_', '{ x }', {}), /^Error: ACCESS_DENIED: not approved/);
  delete props.TOKEN;
  ctx.__fetch = () => resp(400, '{"error":"shop_not_permitted"}');
  assert.throws(() => call('getToken_', true), /FATAL_AUTH.*same Shopify organization/);
});
test('fetchPaged_: batches and resumes after a deadline', () => {
  props.TOKEN = 'x'; props.TOKEN_EXPIRES = String(Date.now() + 86400000);
  const pages = [
    { nodes: Array.from({ length: 60 }, (_, i) => ({ n: i })), pageInfo: { hasNextPage: true, endCursor: 'c1' } },
    { nodes: Array.from({ length: 60 }, (_, i) => ({ n: i })), pageInfo: { hasNextPage: true, endCursor: 'c2' } },
    { nodes: Array.from({ length: 5 }, (_, i) => ({ n: i })), pageInfo: { hasNextPage: false, endCursor: null } }
  ];
  let i = 0;
  ctx.__fetch = () => resp(200, { data: { orders: pages[i++] } });
  const spec = run('ORDER_SPEC'); const seen = [];
  const out = call('fetchPaged_', spec, 'q', null, Date.now() + 60000, b => seen.push(b.length));
  assert.strictEqual(out.done, true); assert.deepStrictEqual(seen, [120, 5]);
  i = 0; const seen2 = [];
  const out2 = call('fetchPaged_', spec, 'q', null, Date.now() - 1, b => seen2.push(b.length));
  assert.deepStrictEqual([out2.done, out2.after, seen2[0]], [false, 'c1', 60]);
});
test('syncSource_: drafts end-to-end into Drafts sheet, query uses updated_at', () => {
  props.TOKEN = 'x'; props.TOKEN_EXPIRES = String(Date.now() + 86400000);
  const sh = new FakeSheet(); sheets['Drafts'] = sh;
  call('ensureHeaders_', sh, DRAFT_HEADERS, run('DRAFT_LAYOUT'));
  let vars;
  ctx.__fetch = (u, o) => { vars = JSON.parse(o.payload).variables; return resp(200, { data: { draftOrders: { nodes: [draft, Object.assign({}, draft, { legacyResourceId: '778', name: '#D13', status: 'COMPLETED', order: { name: '#1060' } })], pageInfo: { hasNextPage: false, endCursor: null } } } }); };
  const src = run('SOURCES').find(s => s.key === 'drafts');
  const out = call('syncSource_', src, cfg, new Date('2026-10-14T00:00:00Z'), null, Date.now() + 60000);
  assert.deepStrictEqual([out.stats.added, out.stats.updated], [2, 0]);
  assert.match(vars.q, /^updated_at:>='2026-10-14T00:00:00Z'$/);
  const m = call('ensureHeaders_', sh, DRAFT_HEADERS);
  assert.strictEqual(sh._cell(3, m['Order #']), '#1060');
  const o = run('SOURCES').find(s => s.key === 'orders');
  ctx.__fetch = (u, x) => { vars = JSON.parse(x.payload).variables; return resp(200, { data: { orders: { nodes: [], pageInfo: { hasNextPage: false } } } }); };
  sheets['Orders'] = new FakeSheet();
  call('syncSource_', o, cfg, new Date('2026-10-14T00:00:00Z'), null, Date.now() + 60000);
  assert.ok(vars.q.endsWith('test:false'));
});
test('resolveRecovered_: token, guess, no match', () => {
  props.TOKEN = 'x'; props.TOKEN_EXPIRES = String(Date.now() + 86400000);
  const sh = new FakeSheet(); sheets['Abandoned checkouts'] = sh;
  const mk = (id, extra) => call('checkoutToRecord_', Object.assign({}, checkout, { id: 'gid://shopify/AbandonedCheckout/' + id, name: '#' + id, completedAt: '2026-10-14T01:10:00Z' }, extra), cfg);
  call('upsert_', sh, ABANDONED_HEADERS, 'Shopify ID', [mk(1, {}), mk(2, { abandonedCheckoutUrl: 'https://s/none' }), mk(3, { abandonedCheckoutUrl: 'https://s/none', customer: { displayName: 'Q', defaultEmailAddress: { emailAddress: 'q@x.co' } } })]);
  ctx.__fetch = (u, o) => {
    const q = JSON.parse(o.payload).variables.q;
    if (q.startsWith('checkout_token:')) return resp(200, { data: { orders: { nodes: [{ name: '#2001' }] } } });
    if (q.includes('s@x.co')) return resp(200, { data: { orders: { nodes: [{ name: '#2002', createdAt: '2026-10-15T00:00:00Z', email: 's@x.co' }] } } });
    return resp(200, { data: { orders: { nodes: [] } } });
  };
  // checkout 1 and 2 both have a token-looking link? only #1 does; #2/#3 links are short
  call('resolveRecovered_', cfg, Date.now() + 60000);
  const m = call('ensureHeaders_', sh, ABANDONED_HEADERS);
  assert.strictEqual(sh._cell(2, m['Order #']), '#2001');
  assert.strictEqual(sh._cell(3, m['Order #']), '#2002 (guess)');
  assert.strictEqual(sh._cell(4, m['Order #']), 'no match');
});

const ID = ORDER_LAYOUT.indexOf('Shopify ID'), REM = ORDER_LAYOUT.indexOf('Remarks');
const mkOrder = (id, iso) => call('orderToRecord_', Object.assign({}, order, { legacyResourceId: String(id), name: '#' + id, createdAt: iso }), cfg);
const ids = sh => sh.data.slice(1, sh.getLastRow()).map(r => r[ID]);
test('upsert_ top: new rows at row 2 newest first, Remarks stay with their rows', () => {
  const sh = new FakeSheet();
  call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  const opts = { mode: 'top', dateHeader: 'Date placed', cfg };
  call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [mkOrder(1, '2026-10-01T00:00:00Z'), mkOrder(2, '2026-10-02T00:00:00Z')], opts);
  assert.deepStrictEqual(ids(sh), ['2', '1']);
  sh._set(3, REM + 1, 'note on 1');
  const paid = Object.assign(mkOrder(1, '2026-10-01T00:00:00Z'), { Status: 'Paid' });
  const r = call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [mkOrder(3, '2026-10-03T00:00:00Z'), paid, mkOrder(4, '2026-10-04T00:00:00Z')], opts);
  assert.deepStrictEqual([r.added, r.updated, r.insertedAtTop], [2, 1, true]);
  assert.deepStrictEqual(ids(sh), ['4', '3', '2', '1']);
  assert.strictEqual(sh._cell(5, REM + 1), 'note on 1');
  assert.strictEqual(sh._cell(5, ORDER_LAYOUT.indexOf('Status') + 1), 'Paid');
});
test('upsert_ sorted: rows interleave by date, bottom-up inserts, blanks last', () => {
  const sh = new FakeSheet();
  call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  const opts = { mode: 'sorted', dateHeader: 'Date placed', cfg };
  call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [mkOrder(10, '2026-10-10T00:00:00Z'), mkOrder(5, '2026-10-05T00:00:00Z'), mkOrder(1, '2026-10-01T00:00:00Z')], opts);
  assert.deepStrictEqual(ids(sh), ['10', '5', '1']);
  sh._set(3, REM + 1, 'five');
  const r = call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [mkOrder(3, '2026-10-03T00:00:00Z'), mkOrder(12, '2026-10-12T00:00:00Z'), mkOrder(0, '2026-09-30T00:00:00Z'), mkOrder(7, '2026-10-07T00:00:00Z'), mkOrder(6, '2026-10-06T00:00:00Z')], opts);
  assert.strictEqual(r.insertedAtTop, true);
  assert.deepStrictEqual(ids(sh), ['12', '10', '7', '6', '5', '3', '1', '0']);
  assert.strictEqual(sh.data.find(x => x[ID] === '5')[REM], 'five');
  // a row with no date: new rows go above it
  sh._set(10, 1, ''); sh._set(10, ID + 1, '99');
  call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [mkOrder(-1, '2026-09-01T00:00:00Z')], opts);
  assert.deepStrictEqual(ids(sh).slice(-2), ['-1', '99']);
});
test('upsert_ without opts still appends at the bottom', () => {
  const sh = new FakeSheet();
  call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [mkOrder(1, '2026-10-01T00:00:00Z')]);
  call('upsert_', sh, ORDER_HEADERS, 'Shopify ID', [mkOrder(2, '2026-10-02T00:00:00Z')]);
  assert.deepStrictEqual(ids(sh), ['1', '2']);
});
const ordersPage = nodes => (u, o) => {
  const q = JSON.parse(o.payload).query;
  const root = q.includes('draftOrders') ? 'draftOrders' : q.includes('abandonedCheckouts') ? 'abandonedCheckouts' : 'orders';
  return resp(200, { data: { [root]: { nodes: root === 'orders' ? nodes : [], pageInfo: { hasNextPage: false, endCursor: null } } } });
};
test('syncSource_ resort: sorts only when rows were added; top re-anchors the status rule', () => {
  props.TOKEN = 'x'; props.TOKEN_EXPIRES = String(Date.now() + 86400000);
  const sh = new FakeSheet(); sheets['Orders'] = sh;
  call('ensureHeaders_', sh, ORDER_HEADERS, ORDER_LAYOUT);
  const src = run('SOURCES').find(s => s.key === 'orders');
  const node = (id, iso) => Object.assign({}, order, { legacyResourceId: String(id), createdAt: iso });
  const rc = Object.assign({}, cfg, { row_order_orders: 'resort' });
  ctx.__fetch = ordersPage([node(1, '2026-10-01T00:00:00Z'), node(3, '2026-10-03T00:00:00Z'), node(2, '2026-10-02T00:00:00Z')]);
  call('syncSource_', src, rc, new Date('2026-10-01T00:00:00Z'), null, Date.now() + 60000);
  assert.deepStrictEqual(ids(sh), ['3', '2', '1']);
  assert.strictEqual(sh.sorted, 1);
  call('syncSource_', src, rc, new Date('2026-10-01T00:00:00Z'), null, Date.now() + 60000);
  assert.strictEqual(sh.sorted, 1, 'no new rows, no sort');
  // top mode: the grey rule shifted to row 5 is replaced by one anchored at row 2
  sh.rules = [{ getBooleanCondition: () => ({ getCriteriaValues: () => ['=$H5="Cancelled"'] }) }];
  ctx.__fetch = ordersPage([node(4, '2026-10-04T00:00:00Z')]);
  call('syncSource_', src, Object.assign({}, cfg, { row_order_orders: 'top' }), new Date('2026-10-01T00:00:00Z'), null, Date.now() + 60000);
  assert.strictEqual(ids(sh)[0], '4');
  assert.deepStrictEqual(sh.rules.map(r => r.getBooleanCondition().getCriteriaValues()[0]), ['=$H2="Cancelled"']);
});
test('getSettings_/ensureSettings_: defaults, dropdowns, invalid values, existing values kept', () => {
  const st = new FakeSheet(); sheets['Settings'] = st;
  st.data = [['country_code', '254'], ['sync_every', 'weekly'], ['sync_hour', '30'], ['row_order_orders', 'sideways']];
  call('ensureSettings_');
  const keys = st.data.map(r => r[0]);
  assert.ok(['row_order_drafts', 'row_order_abandoned', 'timezone'].every(k => keys.includes(k)));
  assert.strictEqual(st.data[0][1], '254', 'existing value untouched');
  assert.strictEqual(keys.filter(k => k === 'sync_every').length, 1, 'no duplicate rows');
  assert.ok(Object.values(st.validation).some(v => v.list.includes('resort')));
  let c = call('getSettings_');
  assert.deepStrictEqual([c.sync_every, c.sync_hour, c.row_order_orders, c.row_order_drafts], ['1h', 6, 'top', 'top']);
  assert.deepStrictEqual([c.sync_interval_ms, c.schedule_id], [3600000, '1h']);
  st.data[1][1] = 'daily'; st.data[2][1] = '7';
  c = call('getSettings_');
  assert.deepStrictEqual([c.sync_interval_ms, c.schedule_id], [86400000, 'daily@7']);
  st.data[1][1] = '15m';
  assert.strictEqual(call('getSettings_').sync_interval_ms, 900000);
});
const trig = () => ctx.ScriptApp.triggers.filter(t => t.handler === 'syncAll');
test('installTrigger_: minutes, hours, daily at hour in Settings timezone', () => {
  ctx.ScriptApp.triggers = []; ctx.__user = 'owner@x.co';
  const base = { timezone: 'Africa/Kampala', sync_hour: 6 };
  call('installTrigger_', Object.assign({}, base, { sync_every: '15m', schedule_id: '15m' }));
  assert.deepStrictEqual(J(trig()[0].calls), [['timeBased'], ['everyMinutes', 15]]);
  call('installTrigger_', Object.assign({}, base, { sync_every: '4h', schedule_id: '4h' }));
  assert.strictEqual(trig().length, 1, 'old trigger removed');
  assert.deepStrictEqual(J(trig()[0].calls), [['timeBased'], ['everyHours', 4]]);
  call('installTrigger_', Object.assign({}, base, { sync_every: 'daily', schedule_id: 'daily@6' }));
  assert.deepStrictEqual(J(trig()[0].calls), [['timeBased'], ['everyDays', 1], ['atHour', 6], ['inTimezone', 'Africa/Kampala']]);
  assert.deepStrictEqual([props.SYNC_SCHEDULE, props.TRIGGER_OWNER], ['daily@6', 'owner@x.co']);
});
test('runSync_: applies a changed schedule, owner only', () => {
  props.TOKEN = 'x'; props.TOKEN_EXPIRES = String(Date.now() + 86400000);
  sheets['Settings'].data = [['sync_every', '2h']];
  ['Orders', 'Drafts', 'Abandoned checkouts'].forEach(n => { sheets[n] = new FakeSheet(); });
  ctx.ScriptApp.triggers = []; ctx.__user = 'owner@x.co'; props.SYNC_SCHEDULE = '1h'; props.TRIGGER_OWNER = 'owner@x.co';
  ctx.__fetch = ordersPage([]);
  call('runSync_');
  assert.strictEqual(props.SYNC_SCHEDULE, '2h');
  assert.deepStrictEqual(J(trig()[0].calls), [['timeBased'], ['everyHours', 2]]);
  sheets['Settings'].data[0][1] = '30m'; ctx.__user = 'someone@else.co';
  call('runSync_');
  assert.strictEqual(props.SYNC_SCHEDULE, '2h', 'non-owner run does not install a second trigger');
  ctx.__user = 'owner@x.co';
});
test('runSync_: resumes from the saved page cursor with the same query, then advances the cursor', () => {
  ['ORDERS', 'DRAFTS', 'ABANDONED'].forEach(k => { props[k + '_CURSOR'] = '2026-10-14T00:00:00.000Z'; });
  props.ORDERS_CURSOR_AFTER = 'c1'; props.ORDERS_CURSOR_SINCE = '2026-10-13T23:50:00.000Z'; props.ORDERS_CURSOR_STARTED = '2026-10-14T05:00:00.000Z';
  const seen = [];
  const page = ordersPage([]);
  ctx.__fetch = (u, o) => { const v = JSON.parse(o.payload); if (!v.query.includes('draftOrders') && !v.query.includes('abandoned')) seen.push(v.variables); return page(u, o); };
  call('runSync_');
  assert.deepStrictEqual([seen[0].after, seen[0].q], ['c1', "updated_at:>='2026-10-13T23:50:00Z' test:false"]);
  assert.strictEqual(props.ORDERS_CURSOR, '2026-10-14T05:00:00.000Z', 'cursor = start of the interrupted run');
  assert.ok(!('ORDERS_CURSOR_AFTER' in props) && !('ORDERS_CURSOR_SINCE' in props));
});
test('runSync_: out of time saves the page cursor and schedules continueSync', () => {
  ctx.ScriptApp.triggers = [];
  let n = 0;
  ctx.__fetch = (u, o) => {
    const q = JSON.parse(o.payload).query;
    if (q.includes('query Orders')) { n++; return resp(200, { data: { orders: { nodes: [], pageInfo: { hasNextPage: true, endCursor: 'p' + n } } } }); }
    return ordersPage([])(u, o);
  };
  const realNow = Date.now; let t = realNow();
  ctx.Date.now = () => (t += 60 * 1000); // each call moves a minute, so the 4.5 min budget runs out
  try { call('runSync_'); } finally { ctx.Date.now = realNow; }
  assert.ok(props.ORDERS_CURSOR_AFTER && props.ORDERS_CURSOR_SINCE && props.ORDERS_CURSOR_STARTED);
  assert.strictEqual(ctx.ScriptApp.triggers.filter(t => t.handler === 'continueSync').length, 1);
  ['_AFTER', '_SINCE', '_STARTED'].forEach(s => delete props['ORDERS_CURSOR' + s]);
});
test('resolveRecovered_: no-match waits at least two sync intervals', () => {
  props.TOKEN = 'x'; props.TOKEN_EXPIRES = String(Date.now() + 86400000);
  const sh = new FakeSheet(); sheets['Abandoned checkouts'] = sh;
  const ago = min => new Date(Date.now() - min * 60000).toISOString();
  const mk = (id, min) => call('checkoutToRecord_', Object.assign({}, checkout, { id: 'gid://shopify/AbandonedCheckout/' + id, abandonedCheckoutUrl: 'https://s/none', completedAt: ago(min), customer: { displayName: 'Q', defaultEmailAddress: { emailAddress: 'q@x.co' } } }), cfg);
  call('upsert_', sh, ABANDONED_HEADERS, 'Shopify ID', [mk(1, 45), mk(2, 180)]);
  ctx.__fetch = () => resp(200, { data: { orders: { nodes: [] } } });
  call('resolveRecovered_', Object.assign({}, cfg, { sync_interval_ms: 3600000 }), Date.now() + 60000);
  const m = call('ensureHeaders_', sh, ABANDONED_HEADERS);
  assert.strictEqual(sh._cell(2, m['Order #']), '');
  assert.strictEqual(sh._cell(3, m['Order #']), 'no match');
});

console.log('\n' + passed + ' passed' + (process.exitCode ? ', with failures' : ''));
