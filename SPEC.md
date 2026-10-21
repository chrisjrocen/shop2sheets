# Shopify → Google Sheets Sync: Orders, Drafts & Abandoned Checkouts — Spec

Status: v3.1 (implemented in `Code.gs`, 15 offline tests passing in `test_code.js`; field names and scopes re-checked against the 2026-10 API docs on 2026-10-14; spec text aligned with the code on 2026-10-16)
Owner: store owner (developer), store owner login, Shopify **Basic** plan

## 1. Problem

The current third-party app forces orders through its own checkout form. That form has two buttons, and disabling one removes both. The owner wants to keep Shopify's **default checkout** and have orders and abandoned checkouts written to Google Sheets automatically, with no dependency on a third-party app.

## 2. Goals / Non-goals

**Goals**
- Three tabs mirroring the Shopify admin menu (Orders → All orders / Drafts / Abandoned checkouts), each row showing date, ID, client name, items, price, phone and location.
- Every online-store order, draft-order-turned-order and manual order appears as one row in an `Orders` tab; every draft order appears in a `Drafts` tab.
- Rows update in place when the order changes (cancelled, refunded, paid, fulfilled, edited).
- Abandoned checkouts with contact details appear in an `Abandoned checkouts` tab and flip to `Recovered` when the customer completes the purchase.
- Free: no paid services, no server. Google Apps Script plus the Shopify Admin GraphQL API.
- Every tab has a **Remarks** column for the owner that the script never writes to, and that stays attached to its row through sorting and filtering.
- Email alert on failure; no duplicate rows ever.

**Non-goals (v1)**
- Two-way sync (sheet edits never go back to Shopify).
- WhatsApp or email notifications for new orders or carts.
- Multiple stores, inventory sync, customer segmentation.
- Real-time updates (up to about 5 minutes of delay is accepted).

## 3. Decisions

| Topic | Decision |
|---|---|
| Spreadsheet | Google Sheets, new spreadsheet with `Orders`, `Drafts`, `Abandoned checkouts`, `Settings`, `Log` tabs. Script is container-bound. |
| Approach | Apps Script only, no hosting |
| Sync method | **Polling** via time-driven trigger (every 5 min) against the Admin GraphQL API |
| Why not webhooks | Apps Script web apps cannot read request headers (no HMAC verification) and answer POSTs with a redirect that Shopify may count as a failed delivery. Polling avoids both and self-heals. |
| Auth | Dev Dashboard app + **client credentials grant** (see 5). Tokens last 24 hours, so the script requests and caches them itself. |
| Scopes | `read_orders` + `read_customers` + `read_draft_orders` (all read-only). `read_customers` is required because an abandoned checkout exposes its email only through `customer`, and the `Customer` object requires that scope. |
| Row granularity | One row per order, items in one cell |
| Order changes | Update the same row, matched by Order ID |
| Order scope | All orders incl. manual and draft-converted orders, cancelled orders kept and marked; **test orders excluded** |
| Abandoned scope | Only checkouts with email or phone; separate tab |
| Recovery | Status from Shopify's own `completedAt`; **order # best-effort** (token match, otherwise labelled guess) |
| Formatting | Phones normalized to `+256…`; times in `Africa/Kampala` |
| API version | Pin `2026-10` (minimum `2026-07`, the first version with `Order.checkoutToken`); review once a year |
| Cost | Free only |

## 4. Verified facts (Shopify docs, checked 2026-10-14)

| Question | Finding | Spec impact |
|---|---|---|
| Does `abandonedCheckouts` need extra scopes? | Docs say it requires `read_orders`. The newest docs also mention the acting *user* needing the `manage_abandoned_checkouts` permission. Third-party guides mention `read_checkouts`, which Shopify's own reference does not. `AbandonedCheckout` has **no top-level email or phone**; contact details come from `customer` (requires `read_customers`) or the shipping/billing address phone. `Customer.email` / `Customer.phone` are deprecated in favour of `defaultEmailAddress { emailAddress }` and `defaultPhoneNumber { phoneNumber }`. | Request `read_orders` + `read_customers` and read the non-deprecated customer fields. If the query still returns an access error, add `read_checkouts` (README troubleshooting). |
| Does the query return recovered checkouts? | Yes: "includes both abandoned and recovered checkouts". The object has `completedAt` (null if not completed), and the query supports `recovery_state:recovered` / `not_recovered`, `status:open/closed`, `updated_at`, `created_at`. | Recovered status comes **directly from Shopify**, with no matching needed. |
| Can an order be linked to its checkout? | `Order.checkoutToken` exists (added to GraphQL in 2026-07) and `orders(query: "checkout_token:…")` works. **`AbandonedCheckout` has no token field.** | Order # matching is best-effort (7.3). |
| Draft and manual orders | `Order.sourceName` (e.g. `web`, `pos`, `shopify_draft_order`), `Order.app.name`, `Order.staffMember` (null for customer-placed orders), and query filter `source_name:`. | `Source` column = `sourceName`; manual orders identifiable by staff member. |
| How to create the app | Since 2026-01-01 new custom apps are created in the **Dev Dashboard** (legacy "Develop apps" creation is closed). No permanent token is shown in the admin. | See 5. |
| Token flow | `POST https://{shop}.myshopify.com/admin/oauth/access_token` with `grant_type=client_credentials`, `client_id`, `client_secret` (form-encoded). Response `expires_in` is always 86399 (24 h). Works only if the **app and the store are in the same Shopify organization**. | Script fetches a token when needed and stores it in **Script Properties** with its expiry, refreshing when under 1 hour is left. (`CacheService` was not used: its maximum lifetime is 6 hours.) Store owner is logging in, so the same-org condition should hold. |
| Protected customer data | A custom app on your own store is reported to get it automatically, but **forum answers from Shopify staff say a merchant-created Dev Dashboard app needs the store on Grow or higher**; other users report Basic working. The docs are not definitive. | **Your plan is Basic, so this is the top risk. See Gate 0 below.** |
| Order history window | Without `read_all_orders`, only orders from the **last 60 days** are readable. | See edge case 11.5. |

### Gate 0: customer-data access test (before building anything else)

Run one query for the most recent real order requesting `phone`, `customer { displayName }`, `shippingAddress { name address1 city }` and `billingAddress { name }` (menu: **Test access**). The same check also confirms `draftOrders` and `abandonedCheckouts` can be read.

- Fields come back populated → proceed as specced.
- Fields null or an `ACCESS_DENIED` / "not approved to access the Customer object" error → customer data is blocked on Basic. Options, in order of preference:
  1. Upgrade to Grow (smallest change; remove the restriction, nothing in the design changes).
  2. Recreate the app under a Shopify Partner organization with **custom distribution** (reported to get protected data automatically). The install uses OAuth instead of client credentials, which is more setup.
  3. Accept an order-data-only version (order #, items, totals, status, no name/phone/address) and take contact details from Shopify admin. Likely not useful for delivery or follow-up, so listed only for completeness.
- This test costs a few minutes and should be done in the first hour of the build, since options 2 and 3 change the design.

## 5. Architecture and Shopify setup

```
Time trigger (every 5 min)
   └─ syncAll()
        ├─ LockService script lock (skip if another run is active)
        ├─ getToken()          // stored in Script Properties, refreshed 1 h before expiry and on 401
        ├─ syncOrders()        // orders updated since cursor - overlap
        ├─ syncDrafts()        // draft orders updated since cursor - overlap
        ├─ syncAbandoned()     // abandoned/recovered checkouts updated since cursor - overlap
        ├─ resolveRecoveredOrders()
        ├─ advance cursors in Script Properties
        └─ on error: Log tab + throttled email
```

Secrets live in **Script Properties** only (never in code or the sheet): `SHOPIFY_STORE`, `CLIENT_ID`, `CLIENT_SECRET`, `ALERT_EMAIL`.

### Setup guide (goes in README)
1. Open the Dev Dashboard (dev.shopify.com/dashboard, or from the store admin: click the store name at the top right → Dev Dashboard).
2. Create an app. Leave the App URL at the default (`https://shopify.dev/apps/default-app-home`). Set scopes `read_orders`, `read_customers` and `read_draft_orders`, release the version, and **install it on your store**.
3. App → Settings → copy **Client ID** and **Client secret**.
4. Confirm the app and your store both appear in the same organization in the Dev Dashboard (a `shop_not_permitted` error means they don't).
5. In the spreadsheet: Extensions → Apps Script → paste code → Project Settings → Script Properties → add the four properties above.
6. Run `setup()` once (authorize), then `testAccess()` (**Gate 0**), then `backfill(60)`.

## 6. Data model

Three tabs, named as in the Shopify admin menu (Orders → All orders, Drafts, Abandoned checkouts). The lean column set below replaces the earlier wide layout (v3). Columns are found **by header name**, so the owner may reorder columns, sort, filter or insert rows. The script writes only the managed columns; **Remarks** is the owner's and is never written, and a hidden **Shopify ID** column is the row key (the visible ID can then be anything). A deleted Remarks column is re-created next to its neighbour; a deleted managed header is re-added and noted in the Log.

### 6.1 `Orders`

| Column | Source |
|---|---|
| Date placed | `createdAt` (Africa/Kampala, `yyyy-mm-dd hh:mm`) |
| Order ID | `name`, e.g. #1042 |
| Client name | shipping name → billing name → customer `displayName` |
| Items ordered | `2x Shirt (M); 1x Cap` |
| Price | **order total** (`totalPriceSet`, presentment; includes shipping and discounts); number format from Settings `price_format` (default `#,##0`) |
| Phone | normalized (section 8) |
| Location | one column: `address1, address2, city` (shipping address, else billing) |
| Status | one column: `Cancelled`, or `<Paid / Unpaid / Partially refunded …> · <Fulfilled / Unfulfilled …>` |
| **Remarks** | owner column, never written |
| Shopify ID | hidden key (`legacyResourceId`) |

Test orders are excluded (`test:false`). Cancelled orders keep their row.

### 6.2 `Drafts` (Shopify `draftOrders`, needs scope `read_draft_orders`)

Date created, Draft ID (`name`), Client name, Items ordered, Price, Phone, Location, **Status** (`Open` / `Invoice sent` / `Completed`), **Order #** (the order created from the draft; kept on the row when the draft is completed), **Remarks**, Shopify ID.

Deleted drafts disappear from the API, but their rows stay in the sheet. Shopify purges drafts after one year of inactivity.

### 6.3 `Abandoned checkouts`

Date started, Checkout ID, Client name, Items ordered, Price, Phone, Email, Location, **Status** (`Abandoned` / `Recovered`), **Order #** (best effort, section 7.3), **Remarks**, Recovered at, Recovery link, Shopify ID. Only checkouts with an email or phone are logged.

### 6.4 Helper tabs
- `Settings`: country code (`256`), timezone, `price_format`, overlap minutes (10), backfill days (60), alert throttle.
- `Log`: timestamp, run type, rows added/updated, errors; trimmed to ~1000 rows.

## 7. Sync behaviour

### 7.1 Orders
- Query `orders(query: "updated_at:>='<cursor − overlap>' test:false", sortKey: UPDATED_AT)` with cursor pagination and a `first` size that respects the query cost budget.
- Build one in-memory map per run: read the hidden Shopify ID column, then `id → row`. Found → overwrite the **managed columns only** on that row (never Remarks). Not found → append.
- Columns are located **by header name**, so the owner may reorder, sort, filter, or insert rows without breaking anything.
- The cursor advances only after the whole run succeeds. The 10-minute overlap plus dedupe protects against boundary misses.

### 7.2 Abandoned checkouts
- Query `abandonedCheckouts(query: "updated_at:>='<cursor − overlap>'")`, paginated. Drop checkouts with neither email nor phone.
- Upsert by Shopify ID. Update items and price if the cart changed.
- `AbandonedCheckoutSortKeys` has no update-time key (only `CHECKOUT_ID`, `CREATED_AT`, `CUSTOMER_NAME`, `ID`, `RELEVANCE`, `TOTAL_PRICE`). The script filters by `updated_at` with the default `ID` sort and pages through all matches; order does not matter because every page is upserted. Shopify notes such queries can be slower, which is acceptable at this volume.
- Contact details: email from `customer.defaultEmailAddress.emailAddress`; phone from `customer.defaultPhoneNumber.phoneNumber`, falling back to the shipping or billing address phone. Name from `customer.displayName`, falling back to the address name.

### 7.3 Recovery
- **Status:** `completedAt != null` (equivalently `recovery_state:recovered`) → `Recovered`, with `Recovered At = completedAt`. This is Shopify's own judgment and needs no heuristics.
- **Recovered order #, best effort:**
  1. Parse a checkout token from `abandonedCheckoutUrl` (**format to be confirmed against the first real abandoned checkout**; the URL pattern is not documented as containing the token). If found, query `orders(query: "checkout_token:<token>")` → fill `Order #`.
  2. Otherwise, find an order with the same email or phone created after the checkout within 7 days → fill `Order #` as `#1050 (guess)`.
  3. Otherwise, after 30 minutes, write `no match`. The `Recovered` status is still shown.
- The recovered-order lookup runs only for newly recovered rows that still lack an order #, so it doesn't re-query every run.

### 7.4 Cancelled / refunded / edited orders
- Any change bumps `updatedAt`, so the next poll refreshes the row. Cancelled orders keep their row with Status = `Cancelled`.
- Owner edits inside managed cells are overwritten on the next update by design. Put notes in Remarks.

### 7.5 Draft and manual orders
- Drafts have their own tab (6.2), synced from `draftOrders(sortKey: UPDATED_AT, query: "updated_at:>=…")`. When a draft is completed its row stays, with Status `Completed` and the new order #. The resulting order also appears on the Orders tab.

## 8. Normalization rules

- **Phone:** strip spaces, dashes and parentheses. `07XXXXXXXX` → `+2567XXXXXXXX`; `2567…` → `+2567…`; already `+` → keep. Invalid or odd numbers are kept as received and highlighted. Write as plain text (cell format `@`).
- **Time:** convert to `Africa/Kampala`, format `yyyy-MM-dd HH:mm`.
- **Money:** write the order total as a number in the order's presentment currency, formatted with Settings `price_format` (default `#,##0`; UGX has no decimals). There is no separate currency column (see 11.12).
- **Items:** `quantity x title (variant)` joined with `; `, skipping variant label "Default Title".

## 9. Reliability

- **Dedupe:** unique keys are Order ID / Checkout ID; repeated or overlapping runs never create duplicates; `LockService` prevents concurrent execution.
- **Backfill:** `backfill(days)` (default 60, max 60 without `read_all_orders`), chunked and resumable via a one-off trigger when nearing the 6-minute limit.
- **Token handling:** store in Script Properties with its expiry and refresh 1 h early; on 401 fetch a new token once and retry; if the token request fails (bad secret, `shop_not_permitted`, app uninstalled), the run fails with a `FATAL_AUTH` message and an alert email (throttled). The trigger keeps running, so syncing resumes by itself once the cause is fixed.
- **Backfill failure:** if a backfill errors, its saved state is cleared so regular syncing is not blocked; re-run Backfill after fixing the cause.
- **Errors:** try/catch around every run; write to `Log`; email `ALERT_EMAIL`, throttled to one per hour per error type.
- **Rate limits:** read `extensions.cost.throttleStatus`; sleep and retry on `THROTTLED`; exponential backoff on 5xx.
- **Quotas:** a 5-minute trigger using a few seconds per run is well within free Apps Script runtime and `UrlFetch` limits.
- **API version:** pinned in one constant; review yearly as Shopify retires old versions.

## 10. UX

- All config lives in `Settings`; no code edits after setup.
- Custom menu **Shopify Sync**: Sync now, Backfill…, Test access (Gate 0), Show last error, Re-install trigger.
- Frozen header, filters on. Conditional formatting greys out finished rows: `Cancelled` orders, `Completed` drafts, `Recovered` checkouts. Guessed recovery matches are marked in the text itself (`#1050 (guess)`).

## 11. Edge cases

1. Order without phone or email (guest, pickup) → blank, not an error.
2. Cash-on-delivery or unpaid orders → included; Status reads `Unpaid · Unfulfilled` until paid.
3. Order edited after creation → Items, Total and Current Total refresh.
4. Partial refund → Status reads `Partially refunded · …`; Price stays the original order total.
5. **Orders older than 60 days:** without `read_all_orders` they cannot be read, so a late refund or cancellation on an old order will not update its row. Accept in v1; apply for `read_all_orders` if it matters.
6. Very long item lists → truncate the cell at ~1000 chars with `…`.
7. Owner deletes a row → it is recreated the next time that order changes; `backfill()` can also restore it.
8. Owner renames or deletes a managed header → the script re-adds it at the end of the managed block and logs a warning.
9. Several abandoned checkouts per customer → each is its own row; each flips to `Recovered` according to Shopify.
10. Customer fields blank/denied (protected data) → detect and report a clear error instead of writing empty rows silently (Gate 0).
11. Checkout with no contact details → not logged; if contact details are added later it appears at the next sync.
12. Multi-currency → Price is in the currency the customer paid in, with no currency column; the store sells in UGX only, so this is accepted. Add a column if other currencies are enabled later.

## 12. Tradeoffs accepted

- Up to ~5 minutes of delay instead of real time.
- Polling uses more API calls than webhooks, which is negligible at this volume.
- Apps Script has a 6-minute execution cap and no real test harness; fine at this scale.
- Recovered order # may be a guess; the Recovered status itself is exact.
- 60-day readable order window unless `read_all_orders` is approved.
- Should volume reach thousands of orders a day, move to a hosted app with webhooks.

## 13. Deliverables

1. `Code.gs`: auth, sync, upsert, normalization, backfill, menu, alerts
2. `appsscript.json`: manifest (timezone `Africa/Kampala`, scopes)
3. `README.md`: setup guide (section 5), Gate 0, troubleshooting (`shop_not_permitted`, access denied, token errors)
4. `test_code.js`: offline tests in Node with a fake spreadsheet and a fake Shopify (`node test_code.js`)
5. Spreadsheet layout created by `setup()` (headers as in section 6)

## 14. Acceptance criteria

- Gate 0 passes: a real order returns name, email, phone and address.
- A new test-gateway order is **not** logged; a real order appears within ~6 minutes with correct name, +256 phone, items, total and Kampala time.
- Cancelling it updates the same row (no new row).
- Running `syncAll()` twice produces zero new rows.
- Text typed in Remarks stays with its order after sort, filter, and later syncs.
- A new draft appears in `Drafts`; completing it marks the row `Completed` with the order #, and the order also appears in `Orders`.
- A checkout abandoned with an email appears in `Abandoned checkouts`; completing it flips it to `Recovered` with the time from Shopify, and `Order #` is filled (or marked `(guess)` / `no match`).
- A checkout without email or phone is not logged.
- A broken client secret triggers exactly one alert email within the throttle window.
- A 60-day backfill completes without duplicates or timeouts.

## 15. Remaining verification (cannot be settled from docs alone)

1. **Gate 0** on your Basic plan (section 4). This is the main unresolved risk.
2. Whether `abandonedCheckoutUrl` contains a usable checkout token. Check with the first real abandoned checkout; fall back to contact matching otherwise.
3. ~~Which `AbandonedCheckoutSortKeys` value sorts by update time.~~ Resolved: there is none; the script filters by `updated_at` and uses the default sort (7.2).
4. Whether the acting app needs anything for the `manage_abandoned_checkouts` permission noted in the newest docs (expected to apply to staff users, not app tokens; the first `abandonedCheckouts` call settles it, and Test access reports it).
5. Whether `read_customers` is granted automatically on your plan; Test access reports a scope error if not.

## 16. Sources

- Shopify GraphQL Admin `AbandonedCheckout` object and `abandonedCheckouts` query (api version 2026-10): https://shopify.dev/docs/api/admin-graphql/latest/objects/AbandonedCheckout
- Shopify GraphQL Admin `Order` object and `orders` query: https://shopify.dev/docs/api/admin-graphql/latest/objects/Order
- Changelog, `checkoutToken` added to `Order` (2026-07): https://shopify.dev/changelog/posts/new-checkouttoken-field-added-to-the-order-object
- Dev Dashboard client credentials grant: https://shopify.dev/docs/apps/build/dev-dashboard/get-api-access-tokens
- Protected customer data: https://shopify.dev/docs/apps/launch/protected-customer-data
- Community threads on custom-app changes and plan limits: https://community.shopify.dev/t/dev-account-cannot-access-personal-information-any-longer/32953, https://community.shopify.dev/t/customer-name-missing-in-api/15665, https://community.shopify.dev/t/enable-protected-customer-data-access-for-a-custom-app-created-in-the-dev-dashboard-no-ui-option-available/35445
