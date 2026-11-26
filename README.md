# Quote Management — Daniel Quotation Portal v2.1

Independent **FastAPI + vanilla HTML/CSS/JS** quotation portal, designed for Vercel. No Lovable dependency; does not use the General Booking project or its database.

## Quick start

- Python 3.11+ and `pip install -r requirements.txt`
- `uvicorn app:app --reload --port 8000`
- Open `http://127.0.0.1:8000`. The public page has an **interactive synthetic-only demo**; margin and USD/IDR FX drive Excel sample exports.

## Login email + PIN (deployment staged)

The UI accepts **email + a six-digit numeric PIN**; the PIN is sent to Supabase Auth as a password over HTTPS. PINs are **never hardcoded, stored in this repository, or returned to the client by the server**. Access to supplier quotations requires a live Supabase Auth session, a verified email address, and a matching server-side `ALLOWED_EMAILS` entry.

**Bootstrap policy:** The owner requested initial PIN `123456`, which is widely guessable. Treat it as a strictly temporary setup credential **only**. When the portal detects a repeated or sequential PIN it requires rotation to a different six-digit PIN before showing the quotation form. This client-side rotation alone is not an authorization boundary; the backend therefore keeps preview/generate **disabled by default** using `ENABLE_REAL_QUOTATIONS=false` until the two approved users' PINs are changed and reviewed.

Provision a **separate** Supabase project (do not reuse General Booking or other MS Statistika projects), disable public signup, and create only approved and confirmed accounts in Auth > Users. Set `SUPABASE_URL`, `SUPABASE_PUBLISHABLE_KEY`, and `ALLOWED_EMAILS` in the Quote Management Vercel project settings, never in GitHub. Initially retain `ENABLE_REAL_QUOTATIONS=false`. Verify server authorization, invalid/unknown email rejection, PIN rotation, unauthorized preview/generate rejection and end-to-end XLSX output. Only then switch `ENABLE_REAL_QUOTATIONS=true` and redeploy.

Six-digit PINs have a limited keyspace: for a stronger production service, prefer longer credentials or an additional factor and enforce Supabase rate limits / bot protection. **Never leave `123456` active as a production PIN.** The live processing feature remains disabled until verified.

## Production architecture

- **DEMO**: synthetic pricing and synthetic sample XLSX/ZIP; no supplier files or real quotation metadata.
- **REAL (disabled until configured)**: Supabase Auth verified server-side, with explicit `ALLOWED_EMAILS` allowlist. Needs a **separate Supabase project**; no secrets in this repository.
- **Original supplier XLSX**: authenticated upload, parsed in-memory, checked against Smartoptics Detailed Quote anchors, pricing gross margin (not markup), anti-double-margin, QA, download DRAFT only.
- **T02 output**: preserves supplier's Smartoptics original workbook; marked DRAFT not supplier-issued.
- **T01 output**: requires original approved CommVerge template, provided via **private upload at generation time**. The original branded master is **not included in this public repository**. Uploads are processed in-memory and not persisted.
- **Advanced**: editable customer and terms metadata only; SKU, descriptions, quantity, publisher identity locked.

No customer-ready quotation should be sent before commercial QA. No automated outbound messages.

## Deploy

1. Import `mstatistika/Quote-Management` into Vercel as a new **separate** project. Framework: FastAPI, repository root.
2. **Preview/demo requires no environment variables**; `/api/status` should say `demo_ready: true` and `real_auth_configured: false`.
3. To enable real mode, create a dedicated Supabase project, provision and email-confirm authorized users, then set Vercel environment variables `SUPABASE_URL`, `SUPABASE_PUBLISHABLE_KEY`, `ALLOWED_EMAILS` (only real authorized addresses).
4. Do **not** commit `.env` files, credentials, customer quotations, or the original branded T01 template. `private/DEMO_MASTER.xlsx` is synthetic and safe for public demo.

## Validation

`python -m unittest discover -s tests -v` for portable parser/pricing/export tests. Tests relying on the non-redistributable T01 master are skipped when that file is not present.

## Limitations

- Source adapter currently supports the original Smartoptics Detailed Quote structure only.
- T01 master must be provided separately and rights to use its brand/template must be confirmed by the issuer.
- Demo is **not** equivalent to successful authenticated end-to-end verification; live testing awaits Supabase and Vercel configuration.
- No built-in pricing/tax assumptions beyond explicitly entered margin and FX.
