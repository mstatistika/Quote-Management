# Quote Management — Daniel Quotation Portal v2.1

Independent **FastAPI + vanilla HTML/CSS/JS** quotation portal, designed for Vercel. No Lovable dependency; does not use the General Booking project or its database.

## Quick start

- Python 3.11+ and `pip install -r requirements.txt`
- `uvicorn app:app --reload --port 8000`
- Open `http://127.0.0.1:8000`. The public page has an **interactive synthetic-only demo**; margin and USD/IDR FX drive Excel sample exports.

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
