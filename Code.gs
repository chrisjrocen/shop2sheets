/**
 * Shopify -> Google Sheets sync: orders, draft orders and abandoned checkouts.
 * Container-bound Apps Script. See README.md for setup and SPEC.md for design.
 *
 * Script Properties required:
 *   SHOPIFY_STORE   e.g. my-store (or my-store.myshopify.com)
 *   CLIENT_ID       Dev Dashboard app Client ID
 *   CLIENT_SECRET   Dev Dashboard app Client secret
 *   ALERT_EMAIL     where failure emails go
 */

const API_VERSION = '2026-10';
const RUN_BUDGET_MS = 4.5 * 60 * 1000; // stay under the 6 minute Apps Script limit
const PAGE_SIZE = 25;

// Tab names match the Shopify admin menu (Orders > All orders / Drafts / Abandoned checkouts).
const SHEET_ORDERS = 'Orders';
const SHEET_DRAFTS = 'Drafts';
const SHEET_ABANDONED = 'Abandoned checkouts';

const KEY_HEADER = 'Shopify ID'; // hidden column the script uses to find each row
const OWNER_COLS = ['Remarks'];  // yours: the script never writes to these

// Full column layouts, left to right. Columns listed in OWNER_COLS are created but never written.
const ORDER_LAYOUT = ['Date placed', 'Order ID', 'Client name', 'Items ordered', 'Price', 'Phone', 'Location', 'Status', 'Remarks', 'Shopify ID'];
const DRAFT_LAYOUT = ['Date created', 'Draft ID', 'Client name', 'Items ordered', 'Price', 'Phone', 'Location', 'Status', 'Order #', 'Remarks', 'Shopify ID'];
const ABANDONED_LAYOUT = ['Date started', 'Checkout ID', 'Client name', 'Items ordered', 'Price', 'Phone', 'Email', 'Location', 'Status', 'Order #', 'Remarks', 'Recovered at', 'Recovery link', 'Shopify ID'];

function managed_(layout) {
  return layout.filter(function (h) { return OWNER_COLS.indexOf(h) === -1; });
}

const ORDER_HEADERS = managed_(ORDER_LAYOUT);
const DRAFT_HEADERS = managed_(DRAFT_LAYOUT);
const ABANDONED_HEADERS = managed_(ABANDONED_LAYOUT);

const DATE_COLS = ['Date placed', 'Date created', 'Date started', 'Recovered at'];
const MONEY_COLS = ['Price'];
// Every other managed column is formatted as plain text so nothing a customer types
// (e.g. a note starting with "=") can be interpreted as a formula.

const SETTINGS_DEFAULTS = {
  country_code: '256',
  timezone: 'Africa/Kampala',
  price_format: '#,##0',
  overlap_minutes: '10',
  backfill_days: '60',
  alert_throttle_minutes: '60',
  sync_every: '1h',
  sync_hour: '6',
  row_order_orders: 'top',
  row_order_drafts: 'top',
  row_order_abandoned: 'top'
};

// Where new rows go: top (row 2, newest first), sorted (by date), resort (append, then sort the tab), bottom.
const ROW_ORDERS = ['top', 'sorted', 'resort', 'bottom'];
// Every interval an Apps Script time trigger supports. `daily` runs at Settings > sync_hour.
const SYNC_EVERY = ['5m', '10m', '15m', '30m', '1h', '2h', '4h', '6h', '8h', '12h', 'daily'];
const SYNC_HOURS = Array.from({ length: 24 }, function (_, h) { return String(h); });

// Settings that get a dropdown in the Settings tab.
const SETTING_CHOICES = {
  sync_every: SYNC_EVERY,
  sync_hour: SYNC_HOURS,
  row_order_orders: ROW_ORDERS,
  row_order_drafts: ROW_ORDERS,
  row_order_abandoned: ROW_ORDERS
};

/* ------------------------------------------------------------------ */
/* GraphQL                                                              */
/* ------------------------------------------------------------------ */

const MONEY = '{ presentmentMoney { amount currencyCode } }';
const ADDRESS = '{ name phone address1 address2 city }';
const ITEMS = 'lineItems(first: 20) { pageInfo { hasNextPage } nodes { name title quantity variantTitle } }';
const ORDER_ITEMS = 'lineItems(first: 20) { pageInfo { hasNextPage } nodes { title quantity variantTitle } }';

const ORDERS_QUERY = `
query Orders($first: Int!, $after: String, $q: String) {
  orders(first: $first, after: $after, query: $q, sortKey: UPDATED_AT) {
    pageInfo { hasNextPage endCursor }
    nodes {
      legacyResourceId name createdAt cancelledAt
      displayFinancialStatus displayFulfillmentStatus
      phone
      customer { displayName }
      shippingAddress ${ADDRESS}
      billingAddress ${ADDRESS}
      totalPriceSet ${MONEY}
      ${ORDER_ITEMS}
    }
  }
}`;

const DRAFTS_QUERY = `
query Drafts($first: Int!, $after: String, $q: String) {
  draftOrders(first: $first, after: $after, query: $q, sortKey: UPDATED_AT) {
    pageInfo { hasNextPage endCursor }
    nodes {
      legacyResourceId name createdAt status completedAt
      order { name }
      phone
      customer { displayName }
      shippingAddress ${ADDRESS}
      billingAddress ${ADDRESS}
      totalPriceSet ${MONEY}
      ${ITEMS}
    }
  }
}`;

const ABANDONED_QUERY = `
query Abandoned($first: Int!, $after: String, $q: String) {
  abandonedCheckouts(first: $first, after: $after, query: $q) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id name createdAt completedAt abandonedCheckoutUrl
      customer {
        displayName
        defaultEmailAddress { emailAddress }
        defaultPhoneNumber { phoneNumber }
      }
      shippingAddress ${ADDRESS}
      billingAddress ${ADDRESS}
      totalPriceSet ${MONEY}
      lineItems(first: 20) { pageInfo { hasNextPage } nodes { title quantity variantTitle } }
    }
  }
}`;

const ORDER_LOOKUP_QUERY = `
query OrderLookup($q: String) {
  orders(first: 10, query: $q, sortKey: CREATED_AT) {
    nodes {
      name createdAt email phone
      shippingAddress { phone }
      billingAddress { phone }
    }
  }
}`;

function shopHost_() {
  const raw = (getProp_('SHOPIFY_STORE') || '').trim()
    .replace(/^https?:\/\//, '').replace(/\/.*$/, '');
  if (!raw) throw new Error('SHOPIFY_STORE script property is not set.');
  return raw.indexOf('.') === -1 ? raw + '.myshopify.com' : raw;
}

function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

function delProp_(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}

/** Client-credentials token (valid 24 h). Cached in Script Properties, refreshed 1 h early. */
function getToken_(force) {
  const token = getProp_('TOKEN');
  const expires = Number(getProp_('TOKEN_EXPIRES') || 0);
  if (!force && token && Date.now() < expires - 60 * 60 * 1000) return token;

  const clientId = getProp_('CLIENT_ID');
  const clientSecret = getProp_('CLIENT_SECRET');
  if (!clientId || !clientSecret) throw new Error('FATAL_AUTH: CLIENT_ID / CLIENT_SECRET script properties are not set.');

  const res = UrlFetchApp.fetch('https://' + shopHost_() + '/admin/oauth/access_token', {
    method: 'post',
    payload: { grant_type: 'client_credentials', client_id: clientId, client_secret: clientSecret },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) {
    let hint = '';
    if (text.indexOf('shop_not_permitted') !== -1) {
      hint = ' The app and the store must be in the same Shopify organization (see README).';
    }
    throw new Error('FATAL_AUTH: token request failed (' + code + '): ' + text.slice(0, 300) + hint);
  }
  const body = JSON.parse(text);
  setProp_('TOKEN', body.access_token);
  setProp_('TOKEN_EXPIRES', String(Date.now() + Number(body.expires_in || 86399) * 1000));
  return body.access_token;
}

/** Runs a GraphQL query with token refresh, throttle handling and retries. */
function gql_(query, variables) {
  const url = 'https://' + shopHost_() + '/admin/api/' + API_VERSION + '/graphql.json';
  let attempt = 0;
  let refreshed = false;
  for (;;) {
    attempt++;
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Shopify-Access-Token': getToken_(false) },
      payload: JSON.stringify({ query: query, variables: variables || {} }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code === 401 && !refreshed) { getToken_(true); refreshed = true; continue; }
    if ((code === 429 || code >= 500) && attempt <= 5) { Utilities.sleep(1000 * Math.pow(2, attempt)); continue; }
    if (code !== 200) throw new Error('Shopify HTTP ' + code + ': ' + text.slice(0, 300));

    const body = JSON.parse(text);
    if (body.errors && body.errors.length) {
      const throttled = body.errors.some(function (e) { return e.extensions && e.extensions.code === 'THROTTLED'; });
      if (throttled && attempt <= 5) { Utilities.sleep(1000 * Math.pow(2, attempt)); continue; }
      const denied = body.errors.filter(function (e) { return e.extensions && e.extensions.code === 'ACCESS_DENIED'; })[0];
      if (denied) throw new Error('ACCESS_DENIED: ' + denied.message);
      throw new Error('GraphQL errors: ' + JSON.stringify(body.errors).slice(0, 500));
    }
    return body.data;
  }
}

/* ------------------------------------------------------------------ */
/* Settings, logging, alerts                                            */
/* ------------------------------------------------------------------ */

function getSettings_() {
  const cfg = {};
  Object.keys(SETTINGS_DEFAULTS).forEach(function (k) { cfg[k] = SETTINGS_DEFAULTS[k]; });
  const sh = SpreadsheetApp.getActive().getSheetByName('Settings');
  if (sh && sh.getLastRow() >= 1) {
    sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
      const k = String(r[0]).trim();
      if (k && r[1] !== '' && Object.prototype.hasOwnProperty.call(SETTINGS_DEFAULTS, k)) cfg[k] = String(r[1]).trim();
    });
  }
  cfg.overlap_minutes = Number(cfg.overlap_minutes) || 10;
  cfg.backfill_days = Number(cfg.backfill_days) || 60;
  cfg.alert_throttle_minutes = Number(cfg.alert_throttle_minutes) || 60;
  ['row_order_orders', 'row_order_drafts', 'row_order_abandoned'].forEach(function (k) {
    if (ROW_ORDERS.indexOf(cfg[k]) === -1) cfg[k] = 'top';
  });
  if (SYNC_EVERY.indexOf(cfg.sync_every) === -1) cfg.sync_every = '1h';
  const hour = parseInt(cfg.sync_hour, 10);
  cfg.sync_hour = hour >= 0 && hour <= 23 ? hour : 6;
  const n = parseInt(cfg.sync_every, 10);
  cfg.sync_interval_ms = cfg.sync_every === 'daily' ? 86400000 : n * (/h$/.test(cfg.sync_every) ? 3600000 : 60000);
  cfg.schedule_id = cfg.sync_every === 'daily' ? 'daily@' + cfg.sync_hour : cfg.sync_every;
  return cfg;
}

/** Creates the Settings tab if needed and appends any missing settings (with dropdowns). Never changes existing values. */
function ensureSettings_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Settings');
  if (!sh) {
    sh = ss.insertSheet('Settings');
    sh.setColumnWidth(1, 200);
  }
  const last = sh.getLastRow();
  const have = last ? sh.getRange(1, 1, last, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  const missing = Object.keys(SETTINGS_DEFAULTS).filter(function (k) { return have.indexOf(k) === -1; });
  if (!missing.length) return;
  const rows = missing.map(function (k) { return [k, SETTINGS_DEFAULTS[k]]; });
  const range = sh.getRange(last + 1, 1, rows.length, 2);
  range.setNumberFormat('@');
  range.setValues(rows);
  missing.forEach(function (k, i) {
    if (!SETTING_CHOICES[k]) return;
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(SETTING_CHOICES[k], true).setAllowInvalid(false).build();
    sh.getRange(last + 1 + i, 2).setDataValidation(rule);
  });
  if (last) log_('setup', 'Added settings: ' + missing.join(', '));
}

function log_(type, message) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Log');
    if (!sh) {
      sh = ss.insertSheet('Log');
      sh.appendRow(['Time', 'Type', 'Message']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), type, String(message).slice(0, 1000)]);
    const rows = sh.getLastRow();
    if (rows > 1100) sh.deleteRows(2, rows - 1000);
  } catch (e) {
    console.error('log failed: ' + e);
  }
}

function sendAlert_(error) {
  const msg = String(error && error.message ? error.message : error);
  setProp_('LAST_ERROR', new Date().toISOString() + ' ' + msg);
  const to = getProp_('ALERT_EMAIL');
  if (!to) return;
  const cfg = getSettings_();
  const key = 'ALERT_' + msg.slice(0, 40).replace(/[^A-Za-z0-9]/g, '_');
  const last = Number(getProp_(key) || 0);
  if (Date.now() - last < cfg.alert_throttle_minutes * 60000) return;
  setProp_(key, String(Date.now()));
  MailApp.sendEmail(to, 'Shopify sync problem',
    'The Shopify -> Google Sheets sync hit an error:\n\n' + msg +
    '\n\nSpreadsheet: ' + SpreadsheetApp.getActive().getUrl() +
    '\n\nYou will not get another email for this error for ' + cfg.alert_throttle_minutes + ' minutes.');
}

/* ------------------------------------------------------------------ */
/* Normalization and record building                                    */
/* ------------------------------------------------------------------ */

function normalizePhone_(raw, countryCode) {
  if (!raw) return '';
  const s = String(raw).trim();
  const digits = s.replace(/[^\d]/g, '');
  if (!digits) return '';
  if (s.charAt(0) === '+') return '+' + digits;
  if (digits.indexOf('00') === 0) return '+' + digits.slice(2);
  if (digits.indexOf(countryCode) === 0 && digits.length >= countryCode.length + 9) return '+' + digits;
  if (digits.charAt(0) === '0') return '+' + countryCode + digits.slice(1);
  if (digits.length === 9) return '+' + countryCode + digits;
  return s; // unknown shape: keep as received
}

function humanize_(enumValue) {
  if (!enumValue) return '';
  const s = String(enumValue).toLowerCase().replace(/_/g, ' ');
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function money_(moneyBag) {
  if (!moneyBag || !moneyBag.presentmentMoney) return '';
  const n = Number(moneyBag.presentmentMoney.amount);
  return isNaN(n) ? '' : n;
}

function itemsSummary_(lineItems) {
  if (!lineItems) return { text: '', qty: 0 };
  let qty = 0;
  const parts = (lineItems.nodes || []).map(function (li) {
    qty += Number(li.quantity) || 0;
    const name = li.name || li.title || 'Item';
    const showVariant = li.variantTitle && li.variantTitle !== 'Default Title' && name.indexOf(li.variantTitle) === -1;
    return li.quantity + 'x ' + name + (showVariant ? ' (' + li.variantTitle + ')' : '');
  });
  let text = parts.join('; ');
  if (lineItems.pageInfo && lineItems.pageInfo.hasNextPage) text += '; ...';
  if (text.length > 1000) text = text.slice(0, 997) + '...';
  return { text: text, qty: qty };
}

/** "address, area, city" from a Shopify address (any part may be missing). */
function locationOf_(addr) {
  if (!addr) return '';
  return [addr.address1, addr.address2, addr.city].filter(Boolean).join(', ');
}

/** First non-empty name: shipping, billing, then the customer record. */
function clientName_(ship, bill, customer) {
  return (ship && ship.name) || (bill && bill.name) || (customer && customer.displayName) || '';
}

const PAYMENT_LABELS = {
  PAID: 'Paid', PENDING: 'Unpaid', AUTHORIZED: 'Authorized', PARTIALLY_PAID: 'Partially paid',
  PARTIALLY_REFUNDED: 'Partially refunded', REFUNDED: 'Refunded', VOIDED: 'Voided', EXPIRED: 'Expired'
};

/** "Cancelled", or "<payment> · <fulfillment>" e.g. "Paid · Unfulfilled". */
function orderStatus_(o) {
  if (o.cancelledAt) return 'Cancelled';
  const pay = PAYMENT_LABELS[o.displayFinancialStatus] || humanize_(o.displayFinancialStatus);
  return [pay, humanize_(o.displayFulfillmentStatus)].filter(Boolean).join(' · ');
}

function orderToRecord_(o, cfg) {
  const ship = o.shippingAddress || {};
  const bill = o.billingAddress || {};
  return {
    'Date placed': new Date(o.createdAt),
    'Order ID': o.name,
    'Client name': clientName_(ship, bill, o.customer),
    'Items ordered': itemsSummary_(o.lineItems).text,
    'Price': money_(o.totalPriceSet),
    'Phone': normalizePhone_(o.phone || ship.phone || bill.phone, cfg.country_code),
    'Location': locationOf_(o.shippingAddress) || locationOf_(o.billingAddress),
    'Status': orderStatus_(o),
    'Shopify ID': String(o.legacyResourceId)
  };
}

function draftToRecord_(d, cfg) {
  const ship = d.shippingAddress || {};
  const bill = d.billingAddress || {};
  return {
    'Date created': new Date(d.createdAt),
    'Draft ID': d.name,
    'Client name': clientName_(ship, bill, d.customer),
    'Items ordered': itemsSummary_(d.lineItems).text,
    'Price': money_(d.totalPriceSet),
    'Phone': normalizePhone_(d.phone || ship.phone || bill.phone, cfg.country_code),
    'Location': locationOf_(d.shippingAddress) || locationOf_(d.billingAddress),
    'Status': humanize_(d.status),
    'Order #': d.order && d.order.name ? d.order.name : '',
    'Shopify ID': String(d.legacyResourceId)
  };
}

/** Returns null when the checkout has neither email nor phone (not logged). */
function checkoutToRecord_(c, cfg) {
  const cust = c.customer || {};
  const ship = c.shippingAddress || {};
  const bill = c.billingAddress || {};
  const email = cust.defaultEmailAddress && cust.defaultEmailAddress.emailAddress || '';
  const phone = normalizePhone_(
    (cust.defaultPhoneNumber && cust.defaultPhoneNumber.phoneNumber) || ship.phone || bill.phone,
    cfg.country_code);
  if (!email && !phone) return null;
  const nameFromCustomer = cust.displayName && cust.displayName !== email && cust.displayName !== phone ? cust.displayName : '';
  const id = String(c.id).split('/').pop();
  const rec = {
    'Date started': new Date(c.createdAt),
    'Checkout ID': c.name || id,
    'Client name': ship.name || bill.name || nameFromCustomer,
    'Items ordered': itemsSummary_(c.lineItems).text,
    'Price': money_(c.totalPriceSet),
    'Phone': phone,
    'Email': email,
    'Location': locationOf_(c.shippingAddress) || locationOf_(c.billingAddress),
    'Status': c.completedAt ? 'Recovered' : 'Abandoned',
    'Recovery link': c.abandonedCheckoutUrl || '',
    'Shopify ID': id
  };
  if (c.completedAt) rec['Recovered at'] = new Date(c.completedAt);
  // 'Order #' is left undefined so updates never overwrite the recovered order number.
  return rec;
}

/* ------------------------------------------------------------------ */
/* Sheet helpers                                                        */
/* ------------------------------------------------------------------ */

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Groups sorted 1-based column numbers into contiguous runs. */
function runsOf_(cols) {
  const sorted = cols.slice().sort(function (a, b) { return a - b; });
  const runs = [];
  sorted.forEach(function (c) {
    const last = runs[runs.length - 1];
    if (last && c === last.start + last.width) last.width++;
    else runs.push({ start: c, width: 1 });
  });
  return runs;
}

function sameValue_(a, b) {
  if (b instanceof Date) return a instanceof Date && a.getTime() === b.getTime();
  return String(a) === String(b);
}

/**
 * Makes sure every managed header exists (matched by name, never by position).
 * A brand-new sheet gets `layout` (managed + owner columns) if given. Missing headers on an
 * existing sheet are inserted after the last managed column, so the owner's columns stay put.
 * Returns { header: column }.
 */
function ensureHeaders_(sheet, headers, layout) {
  const readRow = function () {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); });
  };
  let row = readRow();
  const fresh = row.every(function (v) { return !v; });
  if (fresh) {
    const first = layout || headers;
    sheet.getRange(1, 1, 1, first.length).setValues([first]);
    row = readRow();
  }
  const missing = headers.filter(function (h) { return row.indexOf(h) === -1; });
  if (missing.length) {
    const known = headers.map(function (h) { return row.indexOf(h) + 1; }).filter(function (c) { return c > 0; });
    let after = known.length ? Math.max.apply(null, known) : Math.max(sheet.getLastColumn(), 1);
    missing.forEach(function (h) {
      sheet.insertColumnAfter(after);
      after++;
      sheet.getRange(1, after).setValue(h);
    });
    log_('warn', 'Re-added missing header(s): ' + missing.join(', '));
    row = readRow();
  }
  const map = {};
  headers.forEach(function (h) { map[h] = row.indexOf(h) + 1; });
  return map;
}

/** Re-creates owner columns (like Remarks) next to their neighbour if they were deleted. Never writes data. */
function ensureOwnerColumns_(sheet, layout) {
  layout.forEach(function (h, i) {
    if (OWNER_COLS.indexOf(h) === -1) return;
    const row = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(function (v) { return String(v).trim(); });
    if (row.indexOf(h) !== -1) return;
    let after = 0;
    for (let j = i - 1; j >= 0 && !after; j--) after = row.indexOf(layout[j]) + 1;
    if (!after) after = Math.max(sheet.getLastColumn(), 1);
    sheet.insertColumnAfter(after);
    sheet.getRange(1, after + 1).setValue(h);
  });
}

/**
 * Inserts or updates records (objects keyed by header) in a sheet.
 * Rows are found by key value, so the owner can sort, filter, or insert rows freely.
 * Only managed columns are written; a value of undefined leaves the cell alone.
 * opts = { mode, dateHeader, cfg } decides where new rows go (see ROW_ORDERS); default bottom.
 */
function upsert_(sheet, headers, keyHeader, records, opts) {
  opts = opts || {};
  const mode = opts.mode || 'bottom';
  if (!records.length) return { added: 0, updated: 0, insertedAtTop: false };
  const colMap = ensureHeaders_(sheet, headers);
  const keyCol = colMap[keyHeader];
  const lastRow = Math.max(sheet.getLastRow(), 1);

  const keyToRow = {};
  if (lastRow >= 2) {
    sheet.getRange(2, keyCol, lastRow - 1, 1).getValues().forEach(function (r, i) {
      const k = String(r[0]).trim();
      if (k) keyToRow[k] = i + 2;
    });
  }

  const byKey = {};
  records.forEach(function (r) { byKey[r[keyHeader]] = r; }); // last one wins within a batch
  const updates = [];
  const appends = [];
  Object.keys(byKey).forEach(function (k) { (keyToRow[k] ? updates : appends).push(byKey[k]); });

  let updatedRows = 0;
  if (updates.length) {
    const labelByCol = {};
    headers.forEach(function (h) { labelByCol[colMap[h]] = h; });
    const touched = {};
    runsOf_(headers.map(function (h) { return colMap[h]; })).forEach(function (run) {
      const range = sheet.getRange(2, run.start, lastRow - 1, run.width);
      const data = range.getValues();
      let changed = false;
      updates.forEach(function (rec) {
        const r = keyToRow[rec[keyHeader]] - 2;
        for (let c = 0; c < run.width; c++) {
          const v = rec[labelByCol[run.start + c]];
          if (v === undefined) continue;
          if (!sameValue_(data[r][c], v)) { data[r][c] = v; changed = true; touched[r] = true; }
        }
      });
      if (changed) range.setValues(data);
    });
    updatedRows = Object.keys(touched).length;
  }

  let insertedAtTop = false;
  if (appends.length) {
    const maxCol = Math.max.apply(null, headers.map(function (h) { return colMap[h]; }));
    const toRow = function (rec) {
      const row = new Array(maxCol).fill('');
      headers.forEach(function (h) { if (rec[h] !== undefined) row[colMap[h] - 1] = rec[h]; });
      return row;
    };
    if (mode === 'top' || mode === 'sorted') {
      appends.sort(function (a, b) { return newestFirst_(a[opts.dateHeader], b[opts.dateHeader], a[keyHeader], b[keyHeader]); });
      // Each group of records goes in above `pos` (lastRow + 1 means after the last row).
      const groups = [];
      if (mode === 'top' || lastRow < 2) {
        groups.push({ pos: 2, recs: appends });
      } else {
        const dates = sheet.getRange(2, colMap[opts.dateHeader], lastRow - 1, 1).getValues();
        appends.forEach(function (rec) {
          let pos = lastRow + 1;
          for (let i = 0; i < dates.length; i++) {
            const d = dates[i][0];
            if (!(d instanceof Date) || newestFirst_(rec[opts.dateHeader], d, rec[keyHeader], '') < 0) { pos = i + 2; break; }
          }
          const g = groups[groups.length - 1];
          if (g && g.pos === pos) g.recs.push(rec); else groups.push({ pos: pos, recs: [rec] });
        });
      }
      // Bottom-up, so inserting one group never shifts the position of another.
      groups.sort(function (a, b) { return b.pos - a.pos; }).forEach(function (g) {
        const n = g.recs.length;
        if (g.pos === 2) { sheet.insertRowsAfter(1, n); insertedAtTop = true; }
        else if (g.pos > lastRow) sheet.insertRowsAfter(lastRow, n);
        else sheet.insertRowsBefore(g.pos, n);
        sheet.getRange(g.pos, 1, n, maxCol).setValues(g.recs.map(toRow));
        formatRows_(sheet, colMap, headers, g.pos, n, opts.cfg);
      });
    } else {
      const rows = appends.map(toRow);
      const start = lastRow + 1;
      const need = start + rows.length - 1 - sheet.getMaxRows();
      if (need > 0) sheet.insertRowsAfter(sheet.getMaxRows(), need + 100);
      sheet.getRange(start, 1, rows.length, maxCol).setValues(rows);
    }
  }
  return { added: appends.length, updated: updatedRows, insertedAtTop: insertedAtTop };
}

/** Comparator: newer date first, then higher ID; rows without a date go last. */
function newestFirst_(dateA, dateB, keyA, keyB) {
  const ta = dateA instanceof Date ? dateA.getTime() : -Infinity;
  const tb = dateB instanceof Date ? dateB.getTime() : -Infinity;
  if (ta !== tb) return tb > ta ? 1 : -1;
  const na = Number(keyA), nb = Number(keyB);
  if (!isNaN(na) && !isNaN(nb) && na !== nb) return nb - na;
  return String(keyB).localeCompare(String(keyA));
}

/** Sorts a tab by its date column, newest first. All columns move, so Remarks stay with their rows. */
function sortByDate_(sheet, src) {
  const last = sheet.getLastRow();
  if (last < 3) return;
  const colMap = ensureHeaders_(sheet, src.headers);
  sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).sort([
    { column: colMap[src.dateHeader], ascending: false },
    { column: colMap[KEY_HEADER], ascending: false }
  ]);
}

function formatFor_(header, cfg) {
  if (DATE_COLS.indexOf(header) !== -1) return 'yyyy-mm-dd hh:mm';
  if (MONEY_COLS.indexOf(header) !== -1) return (cfg && cfg.price_format) || SETTINGS_DEFAULTS.price_format;
  return '@';
}

/** Inserted rows can inherit the header's bold/grey styling; reset them to plain data rows. */
function formatRows_(sheet, colMap, headers, startRow, n, cfg) {
  sheet.getRange(startRow, 1, n, sheet.getLastColumn()).setFontWeight('normal').setBackground(null);
  headers.forEach(function (h) { sheet.getRange(startRow, colMap[h], n, 1).setNumberFormat(formatFor_(h, cfg)); });
}

function applyFormats_(sheet, headers, layout, cfg) {
  const colMap = ensureHeaders_(sheet, headers, layout);
  const rows = sheet.getMaxRows();
  headers.forEach(function (h) {
    sheet.getRange(1, colMap[h], rows, 1).setNumberFormat(formatFor_(h, cfg));
  });
  const lastCol = Math.max(sheet.getLastColumn(), layout.length);
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#f1f3f4');
  sheet.setFrozenRows(1);
  if (colMap[KEY_HEADER]) sheet.hideColumns(colMap[KEY_HEADER]);
}

function getSheet_(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

/* ------------------------------------------------------------------ */
/* Paging and the three data sources                                    */
/* ------------------------------------------------------------------ */

const ORDER_SPEC = { query: ORDERS_QUERY, root: 'orders' };
const DRAFT_SPEC = { query: DRAFTS_QUERY, root: 'draftOrders' };
const ABANDONED_SPEC = { query: ABANDONED_QUERY, root: 'abandonedCheckouts' };

const SOURCES = [
  { key: 'orders', sheet: SHEET_ORDERS, layout: ORDER_LAYOUT, headers: ORDER_HEADERS, spec: ORDER_SPEC, cursor: 'ORDERS_CURSOR', extra: 'test:false', toRecord: orderToRecord_,
    dateHeader: 'Date placed', orderSetting: 'row_order_orders', ruleHeader: 'Status', ruleValue: 'Cancelled' },
  { key: 'drafts', sheet: SHEET_DRAFTS, layout: DRAFT_LAYOUT, headers: DRAFT_HEADERS, spec: DRAFT_SPEC, cursor: 'DRAFTS_CURSOR', extra: '', toRecord: draftToRecord_,
    dateHeader: 'Date created', orderSetting: 'row_order_drafts', ruleHeader: 'Status', ruleValue: 'Completed' },
  { key: 'abandoned', sheet: SHEET_ABANDONED, layout: ABANDONED_LAYOUT, headers: ABANDONED_HEADERS, spec: ABANDONED_SPEC, cursor: 'ABANDONED_CURSOR', extra: '', toRecord: checkoutToRecord_,
    dateHeader: 'Date started', orderSetting: 'row_order_abandoned', ruleHeader: 'Status', ruleValue: 'Recovered' }
];
const ROW_RULE_COLOR = '#e0e0e0';

/**
 * Pages through a connection, calling onBatch(nodes) for every ~100 nodes.
 * Returns { done, after }. When done is false, `after` is a cursor that is
 * safe to resume from (everything before it has been passed to onBatch).
 */
function fetchPaged_(spec, queryString, after, deadline, onBatch) {
  let cursor = after || null;
  let batch = [];
  for (;;) {
    const data = gql_(spec.query, { first: PAGE_SIZE, after: cursor, q: queryString });
    const conn = data[spec.root];
    (conn.nodes || []).forEach(function (n) { batch.push(n); });
    const more = conn.pageInfo.hasNextPage;
    cursor = conn.pageInfo.endCursor;
    if (batch.length >= 100 || !more || Date.now() > deadline) {
      if (batch.length) onBatch(batch);
      batch = [];
      if (!more) return { done: true, after: null };
      if (Date.now() > deadline) return { done: false, after: cursor };
    }
  }
}

function isoNoMs_(d) {
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** Fetches everything in `src` updated since `sinceDate` and upserts it into its tab. */
function syncSource_(src, cfg, sinceDate, after, deadline) {
  const sheet = getSheet_(src.sheet);
  if (!sheet) throw new Error('"' + src.sheet + '" tab not found. Run setup() first.');
  const stats = { added: 0, updated: 0 };
  const mode = cfg[src.orderSetting] || 'bottom';
  let insertedAtTop = false;
  const q = "updated_at:>='" + isoNoMs_(sinceDate) + "'" + (src.extra ? ' ' + src.extra : '');
  const res = fetchPaged_(src.spec, q, after, deadline, function (nodes) {
    const recs = nodes.map(function (n) { return src.toRecord(n, cfg); }).filter(Boolean);
    const r = upsert_(sheet, src.headers, KEY_HEADER, recs, { mode: mode, dateHeader: src.dateHeader, cfg: cfg });
    stats.added += r.added; stats.updated += r.updated;
    if (r.insertedAtTop) insertedAtTop = true;
  });
  // Rows inserted above row 2 push the grey status rule down; re-anchor it at row 2.
  if (insertedAtTop) addRowRule_(sheet, src.headers, src.ruleHeader, src.ruleValue, ROW_RULE_COLOR);
  if (mode === 'resort' && stats.added > 0) sortByDate_(sheet, src);
  return { res: res, stats: stats };
}

/* ------------------------------------------------------------------ */
/* Recovery: order number for recovered checkouts (best effort)         */
/* ------------------------------------------------------------------ */

function digitsOnly_(s) { return String(s || '').replace(/[^\d]/g, ''); }

function findRecoveredOrder_(row, colMap, cfg) {
  const cell = function (h) { return row[colMap[h] - 1]; };
  const url = String(cell('Recovery link') || '');
  const m = url.match(/\/checkouts\/(?:cn\/)?([A-Za-z0-9_-]{10,})/);
  if (m) {
    const d = gql_(ORDER_LOOKUP_QUERY, { q: 'checkout_token:' + m[1] });
    const hit = d.orders.nodes[0];
    if (hit) return { name: hit.name, method: 'token' };
  }

  const created = cell('Date started');
  if (!(created instanceof Date)) return null;
  const windowEnd = created.getTime() + 7 * 86400000;
  const email = String(cell('Email') || '').trim();
  const phone = String(cell('Phone') || '').trim();
  const base = "created_at:>='" + isoNoMs_(created) + "' test:false";
  let q = null;
  if (email) q = 'email:"' + email.replace(/"/g, '') + '" ' + base;
  else if (phone) q = '"' + phone.replace(/"/g, '') + '" ' + base;
  if (!q) return null;

  const nodes = gql_(ORDER_LOOKUP_QUERY, { q: q }).orders.nodes
    .filter(function (n) { return new Date(n.createdAt).getTime() <= windowEnd; })
    .filter(function (n) {
      if (email) return String(n.email || '').toLowerCase() === email.toLowerCase();
      const want = digitsOnly_(phone);
      return [n.phone, n.shippingAddress && n.shippingAddress.phone, n.billingAddress && n.billingAddress.phone]
        .some(function (p) { return p && digitsOnly_(normalizePhone_(p, cfg.country_code)) === want; });
    })
    .sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); });
  return nodes[0] ? { name: nodes[0].name, method: 'contact guess' } : null;
}

/** Fills "Order #" for recovered checkouts. A match by contact (not by token) is marked "(guess)". */
function resolveRecovered_(cfg, deadline) {
  const sheet = getSheet_(SHEET_ABANDONED);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const colMap = ensureHeaders_(sheet, ABANDONED_HEADERS);
  const width = Math.max.apply(null, ABANDONED_HEADERS.map(function (h) { return colMap[h]; }));
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  let handled = 0;
  for (let i = 0; i < data.length && handled < 15 && Date.now() < deadline; i++) {
    const row = data[i];
    if (row[colMap['Status'] - 1] !== 'Recovered') continue;
    if (String(row[colMap['Order #'] - 1]) !== '') continue;
    handled++;
    const match = findRecoveredOrder_(row, colMap, cfg);
    if (match) {
      sheet.getRange(i + 2, colMap['Order #']).setValue(match.name + (match.method === 'token' ? '' : ' (guess)'));
    } else {
      const at = row[colMap['Recovered at'] - 1];
      // Give every recovery at least one retry before giving up, however long the sync interval is.
      const wait = Math.max(30 * 60000, 2 * (cfg.sync_interval_ms || 0));
      const recent = at instanceof Date && Date.now() - at.getTime() < wait;
      if (!recent) sheet.getRange(i + 2, colMap['Order #']).setValue('no match');
    }
  }
  return handled;
}

/* ------------------------------------------------------------------ */
/* Entry points                                                         */
/* ------------------------------------------------------------------ */

/** Runs on the schedule in Settings > sync_every (time-driven trigger). */
function syncAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) { log_('skip', 'Another run is still active.'); return; }
  try {
    if (getProp_('BF_STAGE')) { log_('skip', 'Backfill in progress.'); return; }
    runSync_();
  } catch (e) {
    handleError_(e, 'sync');
  } finally {
    lock.releaseLock();
  }
}

function handleError_(e, context) {
  const msg = String(e && e.message ? e.message : e);
  log_('error', context + ': ' + msg);
  try { sendAlert_(e); } catch (mailErr) { log_('error', 'alert failed: ' + mailErr); }
}

/** Target of the one-off trigger set when a sync runs out of time. */
function continueSync() {
  removeTriggers_('continueSync');
  syncAll();
}

function runSync_() {
  ensureSettings_();
  const cfg = getSettings_();
  const startedAt = new Date();
  const deadline = Date.now() + RUN_BUDGET_MS;
  const overlap = cfg.overlap_minutes * 60000;
  const errors = [];
  const summary = [];
  const defaultSince = new Date(Date.now() - cfg.backfill_days * 86400000);
  let unfinished = false;

  // Apply a changed Settings > sync_every. Only the trigger's owner may, or a second trigger would appear.
  const owner = getProp_('TRIGGER_OWNER');
  if (getProp_('SYNC_SCHEDULE') !== cfg.schedule_id && (!owner || owner === currentUser_())) {
    try { installTrigger_(cfg); } catch (e) { errors.push('trigger: ' + (e && e.message ? e.message : e)); }
  }

  SOURCES.forEach(function (src) {
    // A run that ran out of time saved its page cursor; resume the same query from there.
    const pending = getProp_(src.cursor + '_AFTER');
    try {
      let since, started;
      if (pending) {
        since = new Date(getProp_(src.cursor + '_SINCE'));
        started = getProp_(src.cursor + '_STARTED') || startedAt.toISOString();
      } else {
        const saved = getProp_(src.cursor);
        since = new Date((saved ? new Date(saved) : defaultSince).getTime() - overlap);
        started = startedAt.toISOString();
      }
      const out = syncSource_(src, cfg, since, pending || null, deadline);
      if (out.res.done) {
        setProp_(src.cursor, started);
        clearSyncResume_(src);
      } else {
        setProp_(src.cursor + '_AFTER', out.res.after);
        setProp_(src.cursor + '_SINCE', since.toISOString());
        setProp_(src.cursor + '_STARTED', started);
        unfinished = true;
        log_('warn', src.key + ': time budget reached, continuing in about a minute.');
      }
      summary.push(src.key + ' +' + out.stats.added + ' ~' + out.stats.updated);
    } catch (e) {
      if (pending) clearSyncResume_(src); // a stale page cursor must not block every later run
      errors.push(src.key + ': ' + (e && e.message ? e.message : e));
    }
  });
  if (unfinished) scheduleSyncContinuation_();

  try {
    const n = resolveRecovered_(cfg, deadline);
    if (n) summary.push('recovery lookups ' + n);
  } catch (e) {
    errors.push('recovery: ' + (e && e.message ? e.message : e));
  }

  log_('sync', summary.join(' | '));
  if (errors.length) throw new Error(errors.join(' || '));
}

/** Imports the last N days (default from Settings). Resumes itself if it runs out of time. */
function backfill(days) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) { log_('skip', 'Another run is active; try again in a minute.'); return; }
  try {
    const cfg = getSettings_();
    const n = Number(days) || cfg.backfill_days;
    setProp_('BF_START', new Date().toISOString());
    setProp_('BF_SINCE', new Date(Date.now() - n * 86400000).toISOString());
    setProp_('BF_STAGE', 'running');
    setProp_('BF_INDEX', '0');
    delProp_('BF_AFTER');
    SOURCES.forEach(clearSyncResume_); // backfill re-reads everything and resets the cursors
    log_('backfill', 'Started: last ' + n + ' days.');
    continueBackfillLocked_();
  } catch (e) {
    clearBackfillState_();
    handleError_(e, 'backfill');
  } finally {
    lock.releaseLock();
  }
}

/** Target of the one-off continuation trigger. */
function continueBackfill() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) { scheduleBackfillContinuation_(); return; }
  try {
    removeTriggers_('continueBackfill');
    if (getProp_('BF_STAGE')) continueBackfillLocked_();
  } catch (e) {
    clearBackfillState_();
    handleError_(e, 'backfill');
  } finally {
    lock.releaseLock();
  }
}

function continueBackfillLocked_() {
  const cfg = getSettings_();
  const deadline = Date.now() + RUN_BUDGET_MS;
  const since = new Date(getProp_('BF_SINCE'));
  let idx = Number(getProp_('BF_INDEX') || 0);

  while (idx < SOURCES.length) {
    const src = SOURCES[idx];
    const out = syncSource_(src, cfg, since, getProp_('BF_AFTER') || null, deadline);
    log_('backfill', src.key + ' +' + out.stats.added + ' ~' + out.stats.updated);
    if (!out.res.done) {
      setProp_('BF_AFTER', out.res.after);
      scheduleBackfillContinuation_();
      return;
    }
    delProp_('BF_AFTER');
    idx++;
    setProp_('BF_INDEX', String(idx));
  }

  const start = getProp_('BF_START');
  SOURCES.forEach(function (src) { setProp_(src.cursor, start); });
  clearBackfillState_();
  log_('backfill', 'Finished.');
}

/** Drops the saved backfill state so regular syncing is not blocked. After an error, re-run Backfill once the cause is fixed. */
function clearBackfillState_() {
  ['BF_START', 'BF_SINCE', 'BF_STAGE', 'BF_AFTER', 'BF_INDEX'].forEach(delProp_);
  removeTriggers_('continueBackfill');
}

function scheduleBackfillContinuation_() {
  removeTriggers_('continueBackfill');
  ScriptApp.newTrigger('continueBackfill').timeBased().after(60 * 1000).create();
}

function clearSyncResume_(src) {
  ['_AFTER', '_SINCE', '_STARTED'].forEach(function (s) { delProp_(src.cursor + s); });
}

function scheduleSyncContinuation_() {
  removeTriggers_('continueSync');
  ScriptApp.newTrigger('continueSync').timeBased().after(60 * 1000).create();
}

function removeTriggers_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

function currentUser_() {
  try { return Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function describeSchedule_(cfg) {
  if (cfg.sync_every === 'daily') return 'daily at ' + cfg.sync_hour + ':00 (' + cfg.timezone + ')';
  const n = parseInt(cfg.sync_every, 10);
  return /h$/.test(cfg.sync_every) ? 'every ' + n + (n === 1 ? ' hour' : ' hours') : 'every ' + n + ' minutes';
}

/** Replaces the current user's sync trigger with one matching Settings > sync_every / sync_hour. */
function installTrigger_(cfg) {
  removeTriggers_('syncAll');
  const n = parseInt(cfg.sync_every, 10);
  const b = ScriptApp.newTrigger('syncAll').timeBased();
  // The manifest timezone is fixed, so the daily hour is pinned to the Settings timezone explicitly.
  if (cfg.sync_every === 'daily') b.everyDays(1).atHour(cfg.sync_hour).inTimezone(cfg.timezone).create();
  else if (/h$/.test(cfg.sync_every)) b.everyHours(n).create();
  else b.everyMinutes(n).create();
  setProp_('SYNC_SCHEDULE', cfg.schedule_id);
  setProp_('TRIGGER_OWNER', currentUser_());
  log_('setup', 'Installed sync trigger: ' + describeSchedule_(cfg) + '.');
}

/**
 * Triggers belong to whoever installs them, and a script cannot delete another user's trigger.
 * Returns false if someone else owns the schedule and the current user declines to add a second one.
 */
function confirmTriggerOwner_() {
  const owner = getProp_('TRIGGER_OWNER');
  const me = currentUser_();
  if (!owner || !me || owner === me) return true;
  const text = 'The sync schedule was installed by ' + owner + '. Installing yours too makes every sync run twice.';
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {
    // Run from the script editor: no dialog possible, so proceed and leave a note.
    log_('warn', text + ' Ask ' + owner + ' to use Shopify Sync > Remove my trigger.');
    return true;
  }
  return ui.alert('Sync schedule', text + '\n\nContinue? Afterwards ask ' + owner +
    ' to use Shopify Sync > Remove my trigger (or delete it in Apps Script > Triggers).',
    ui.ButtonSet.YES_NO) === ui.Button.YES;
}

/** Menu: applies Settings > sync_every now instead of on the next run. */
function installTrigger() {
  ensureSettings_();
  const cfg = getSettings_();
  if (!confirmTriggerOwner_()) return;
  installTrigger_(cfg);
  notify_('Sync schedule', 'Syncing ' + describeSchedule_(cfg) + '.');
}

/** Menu: removes the current user's sync trigger (only your own; Google does not allow removing someone else's). */
function removeMyTrigger() {
  removeTriggers_('syncAll');
  removeTriggers_('continueSync');
  if (getProp_('TRIGGER_OWNER') === currentUser_()) delProp_('TRIGGER_OWNER');
  notify_('Sync schedule', 'Your sync trigger was removed. Use Re-install trigger to sync on a schedule again.');
}

/* ------------------------------------------------------------------ */
/* Setup, menu, Gate 0 test                                             */
/* ------------------------------------------------------------------ */

function setup() {
  const ss = SpreadsheetApp.getActive();
  const missing = ['SHOPIFY_STORE', 'CLIENT_ID', 'CLIENT_SECRET', 'ALERT_EMAIL'].filter(function (k) { return !getProp_(k); });

  ensureSettings_();
  const cfg = getSettings_();
  ss.setSpreadsheetTimeZone(cfg.timezone);

  const firstSheet = ss.getSheets()[0];
  SOURCES.forEach(function (p, i) {
    let sheet = ss.getSheetByName(p.sheet);
    if (!sheet) {
      // Reuse an untouched default sheet instead of leaving an empty "Sheet1" behind.
      if (i === 0 && firstSheet.getName() === 'Sheet1' && firstSheet.getLastRow() === 0) { sheet = firstSheet; sheet.setName(p.sheet); }
      else sheet = ss.insertSheet(p.sheet);
    }
    ensureHeaders_(sheet, p.headers, p.layout);
    ensureOwnerColumns_(sheet, p.layout);
    applyFormats_(sheet, p.headers, p.layout, cfg);
    try {
      if (!sheet.getFilter()) sheet.getRange(1, 1, sheet.getMaxRows(), Math.max(sheet.getLastColumn(), p.layout.length)).createFilter();
    } catch (e) { /* filter is optional */ }
    addRowRule_(sheet, p.headers, p.ruleHeader, p.ruleValue, ROW_RULE_COLOR);
  });
  if (!ss.getSheetByName('Log')) log_('setup', 'Created.');
  const installed = confirmTriggerOwner_();
  if (installed) installTrigger_(cfg);

  const ready = installed
    ? 'Tabs are ready; syncing ' + describeSchedule_(cfg) + ' (change it in Settings > sync_every).'
    : 'Tabs are ready; the sync schedule was left with ' + getProp_('TRIGGER_OWNER') + '.';
  const msg = missing.length
    ? ready + '\n\nThese Script Properties are still missing: ' + missing.join(', ') + '.\n\nAdd them in Project Settings, then run Test access.'
    : ready + '\n\nNext: run Shopify Sync > Test access, then Shopify Sync > Backfill.';
  notify_('Setup finished', msg);
}

function addRowRule_(sheet, headers, header, value, color) {
  const col = ensureHeaders_(sheet, headers)[header];
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const formula = '=$' + colLetter_(col) + '2="' + value + '"';
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(color)
    .setRanges([sheet.getRange(2, 1, sheet.getMaxRows() - 1, lastCol)])
    .build();
  // Also drops our own rule after inserted rows shifted it to start below row 2 (=$H5="Cancelled").
  const shifted = new RegExp('^=\\$' + colLetter_(col) + '\\d+="' + value + '"$');
  const rules = sheet.getConditionalFormatRules().filter(function (r) {
    const c = r.getBooleanCondition();
    return !(c && c.getCriteriaValues && shifted.test(String(c.getCriteriaValues()[0])));
  });
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function notify_(title, message) {
  log_('info', title + ': ' + message);
  try { SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { console.log(title + ': ' + message); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Shopify Sync')
    .addItem('Sync now', 'syncNowMenu_')
    .addItem('Backfill...', 'backfillPrompt_')
    .addItem('Test access (Gate 0)', 'testAccess')
    .addItem('Sort newest first', 'sortNewestFirstMenu_')
    .addItem('Show last error', 'showLastError')
    .addItem('Re-install trigger', 'installTrigger')
    .addItem('Remove my trigger', 'removeMyTrigger')
    .addToUi();
}

/** Menu: one-time sort of every tab by date, newest first. */
function sortNewestFirstMenu_() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('Sort newest first',
    'Sorts Orders, Drafts and Abandoned checkouts by date, newest first. Remarks move with their rows. Continue?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) { notify_('Sort newest first', 'A sync is running; try again in a minute.'); return; }
  try {
    SOURCES.forEach(function (src) { const sh = getSheet_(src.sheet); if (sh) sortByDate_(sh, src); });
  } finally {
    lock.releaseLock();
  }
  notify_('Sort newest first', 'Done.');
}

function syncNowMenu_() { syncAll(); notify_('Sync', 'Finished. See the Log tab for details.'); }

function backfillPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Backfill', 'How many days back? (Shopify only exposes the last 60 days of orders by default)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  backfill(Number(r.getResponseText()) || undefined);
  notify_('Backfill', 'Started. If it is large it continues automatically; watch the Log tab.');
}

function showLastError() {
  notify_('Last error', getProp_('LAST_ERROR') || 'No error recorded.');
}

/**
 * Gate 0: confirms the token works, that customer fields (name, phone, address)
 * are actually returned for your plan, and that drafts and abandoned checkouts are readable.
 * Run this before anything else.
 */
function testAccess() {
  const lines = [];
  try {
    getToken_(true);
    lines.push('OK  token request succeeded.');
  } catch (e) {
    notify_('Test access', 'FAIL  ' + e.message);
    return;
  }

  try {
    const d = gql_(`{ orders(first: 1, query: "test:false", sortKey: CREATED_AT, reverse: true) {
      nodes { name phone customer { displayName } shippingAddress { name address1 city } billingAddress { name } } } }`);
    const o = d.orders.nodes[0];
    if (!o) {
      lines.push('--  no orders found yet; place a real order and run this again.');
    } else {
      const ship = o.shippingAddress || {};
      const filled = [o.phone, ship.name, ship.address1, (o.customer || {}).displayName, (o.billingAddress || {}).name].filter(Boolean).length;
      lines.push(filled
        ? 'OK  order ' + o.name + ' returned customer fields (' + filled + ' of 5 populated).'
        : 'WARN order ' + o.name + ' came back with ALL customer fields empty. Customer data may be blocked on your plan (see README, "If Gate 0 fails").');
    }
  } catch (e) {
    lines.push((String(e.message).indexOf('ACCESS_DENIED') === 0 ? 'FAIL customer data denied: ' : 'FAIL orders query: ') + e.message);
  }

  try {
    const d = gql_(`{ draftOrders(first: 1) { nodes { name status } } }`);
    lines.push('OK  draft orders readable (' + d.draftOrders.nodes.length + ' returned).');
  } catch (e) {
    lines.push('FAIL draft orders: ' + e.message + ' (the app needs the read_draft_orders scope).');
  }

  try {
    const d = gql_(`{ abandonedCheckouts(first: 1) { nodes { id customer { displayName defaultEmailAddress { emailAddress } } } } }`);
    lines.push('OK  abandoned checkouts readable (' + d.abandonedCheckouts.nodes.length + ' returned).');
  } catch (e) {
    lines.push('FAIL abandoned checkouts: ' + e.message + ' (the app needs read_orders and read_customers; if Shopify still refuses, try adding read_checkouts).');
  }
  notify_('Test access', lines.join('\n'));
}
