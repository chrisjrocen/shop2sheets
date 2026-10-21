# Shopify → Google Sheets sync

Writes your Shopify orders and abandoned checkouts into a Google Sheet automatically, using Shopify's normal checkout. Free, no server: one Google Apps Script that checks Shopify every 5 minutes. Design details are in `SPEC.md`.

**What you get** (three tabs, named like the Shopify admin menu)
- `Orders`: Date placed, Order ID, Client name, Items ordered, Price (order total), Phone, Location (address, area, city), Status (e.g. `Paid · Unfulfilled`, or `Cancelled`), **Remarks**.
- `Drafts`: same, plus Order # once the draft is completed (the row stays, marked `Completed`).
- `Abandoned checkouts`: only checkouts with an email or phone. Status flips to `Recovered` and Order # is filled when the customer buys.
- **Remarks** is yours: the script never writes to it. Rows are matched by a hidden `Shopify ID` column, so you can sort, filter and reorder columns safely.
- `Settings` (country code, timezone, price format, overlap, backfill days, alert throttle) and `Log` tabs.

Test orders are skipped.

Expect up to about 5 minutes between an order and its row.

## Setup (about 20 minutes)

### 1. Create the Shopify app
1. Open the Dev Dashboard: <https://dev.shopify.com/dashboard>, or from your store admin click your store name (top right) and choose **Dev Dashboard**. Shopify closed the old "Develop apps → Create an app" flow on January 1, 2026.
2. **Create app**. Name it anything, e.g. `Sheets sync`. Leave the App URL at its default (`https://shopify.dev/apps/default-app-home`).
3. Set the Admin API scopes to **`read_orders`**, **`read_customers`** and **`read_draft_orders`** (read-only; the script never changes anything in your store). `read_customers` is needed because Shopify only exposes an abandoned checkout's email through the customer record; `read_draft_orders` is for the Drafts tab. If you created the app earlier, add the new scope, release a new version, and approve the update on your store.
4. **Release** the version, then **install the app on your store**.
5. Open the app's **Settings** and copy the **Client ID** and **Client secret**.
6. In the Dev Dashboard, confirm that both the app and your store appear under the same organization. If they don't, step 4 of the checks below (`shop_not_permitted`) will fail.

### 2. Create the spreadsheet and paste the code
1. Create a **new** Google Sheet (tab names and columns changed from the first version; if you reuse the old sheet, delete its old `Orders` and `Abandoned` tabs first).
2. **Extensions → Apps Script**. Replace the contents of `Code.gs` with this repo's `Code.gs`.
3. Project Settings (gear icon) → tick **Show "appsscript.json" manifest file in editor**, then open `appsscript.json` and replace it with this repo's version.
4. Project Settings → **Script Properties** → add four properties:

| Property | Value |
|---|---|
| `SHOPIFY_STORE` | your store handle, e.g. `my-store` (the part before `.myshopify.com`) |
| `CLIENT_ID` | Client ID from step 1.5 |
| `CLIENT_SECRET` | Client secret from step 1.5 |
| `ALERT_EMAIL` | where failure emails should go |

### 3. Run it
1. In the Apps Script editor, pick `setup` in the function dropdown and **Run**. Approve the permissions. This creates the tabs, formats the columns, and installs the 5-minute trigger.
2. Back in the sheet, reload the page. A **Shopify Sync** menu appears.
3. **Shopify Sync → Test access (Gate 0).** Read the result before going further (see below).
4. **Shopify Sync → Backfill…** and enter `60`. This imports the last 60 days. If it is large it continues by itself.

From then on the trigger does everything.

## Gate 0: does your plan return customer details?

Test access runs these checks (it also confirms the Drafts and Abandoned checkouts queries work):

| Line | Meaning |
|---|---|
| `OK token request succeeded` | Client ID/secret and the store handle are right |
| `OK order … returned customer fields` | Name, email, phone and address come through |
| `WARN … ALL customer fields empty` or `FAIL customer data denied` | Shopify is withholding customer data from the app |

Your store is on the **Basic** plan, and Shopify's documentation is not clear on whether an app you made yourself can read customer data there. Some Shopify staff replies say Grow or higher is required; other developers report it working on Basic. This is the one thing the code cannot settle, so run Gate 0 first.

### If Gate 0 fails
1. **Upgrade to Grow.** Smallest change; nothing else in the setup changes.
2. **Create the app under a Shopify Partner organization with custom distribution.** Reported to get customer data automatically, but installs through OAuth instead of the simple client-credentials flow used here, so the token code would need to change.
3. **Order data only.** The script still works, but name, phone and address stay blank, which is probably not useful for delivery or follow-up.

## Everyday use

| Menu item | What it does |
|---|---|
| Sync now | Runs one sync immediately |
| Backfill… | Re-imports N days (safe to repeat; never duplicates) |
| Test access (Gate 0) | Checks token, customer data, drafts and abandoned checkouts |
| Show last error | The most recent failure message |
| Re-install trigger | Recreates the 5-minute trigger if it was removed |

The `Log` tab records each run. A failure also emails `ALERT_EMAIL`, at most once per hour per kind of error.

**Abandoned checkouts.** "Recovered" always comes from Shopify. `Order #` is best effort: first by the checkout token if one can be read from the recovery link, otherwise by matching the same email (or phone) on an order within 7 days, shown as `#1050 (guess)`. If neither works it says `no match`.

**Drafts.** Completed drafts keep their row and show the order number. If you delete a draft in Shopify, its row stays in the sheet (Shopify no longer reports it); Shopify itself purges drafts after a year of inactivity.

**Price** is the order total (shipping and discounts included), shown with the format in Settings `price_format` (default `#,##0`).

## Things worth knowing
- **60-day window.** Shopify only lets this app read orders from the last 60 days. An order older than that will not update if it is refunded or cancelled later.
- **Edits in managed columns are overwritten** the next time that order changes. Put notes in **Remarks**.
- **If you delete a row**, it comes back the next time that order changes (or on backfill). Don't delete the hidden `Shopify ID` column; if you do, the script re-adds it and may re-add rows.
- **If you rename or delete a header** in the managed block, the script re-adds it and notes it in the `Log`.
- **API version** is pinned in `Code.gs` (`API_VERSION`). Shopify retires old versions after about a year, so bump it once a year and re-run Test access.

## Troubleshooting

| Message | Cause and fix |
|---|---|
| `FATAL_AUTH … shop_not_permitted` | The app and store are not in the same Shopify organization. Check both appear in your Dev Dashboard (setup step 1.6) |
| `FATAL_AUTH … CLIENT_ID / CLIENT_SECRET … not set` | Script Properties missing or misspelled |
| `FATAL_AUTH … 400/401` | Wrong Client secret, or the app is not installed on the store |
| `ACCESS_DENIED … Customer object` | See Gate 0 above |
| `ACCESS_DENIED` mentioning a scope | Add the missing scope in the app (`read_orders`, `read_customers`, `read_draft_orders`), release a new version, and reinstall |
| `abandoned checkouts query` fails | Check both scopes. If Shopify still asks for it, add `read_checkouts`. The newest docs also mention an abandoned-checkout permission for staff users; the first call tells you whether it applies to the app |
| `"Orders" tab not found` (or Drafts / Abandoned checkouts) | Run `setup()` |
| Nothing new appears | Check **Triggers** (clock icon in Apps Script) for a `syncAll` trigger; use *Re-install trigger* |

## Tests

`test_code.js` runs `Code.gs` in Node with a fake spreadsheet and a fake Shopify, no accounts needed:

```
node test_code.js
```

It covers phone formatting, building order/draft/checkout rows, header matching, Remarks surviving sorting, never overwriting the recovered Order #, the Drafts sync, recovered-order matching, token refresh on 401, throttle handling, and resuming a paged fetch. It cannot test real Shopify responses; Gate 0 and the first real order are the true end-to-end checks.
