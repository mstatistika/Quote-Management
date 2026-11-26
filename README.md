# Quote Generator — v2.3.1

Independent **FastAPI + vanilla HTML/CSS/JS** quotation portal, designed for Vercel. No Lovable dependency; does not use the General Booking project or its database.

## Quick start

- Python 3.11+ and `pip install -r requirements.txt`
- `uvicorn app:app --reload --port 8000`
- Open `http://127.0.0.1:8000`. The public page has an **interactive synthetic-only demo**; margin and USD/IDR FX drive Excel sample exports.

## Login pengguna

Select Michael or Daniel from the email dropdown, enter the account's existing password, then sign in. No mandatory password/PIN rotation and no OTP in the MVP. Future password management can be added to a user menu.

Credentials are **not** embedded in source code. The browser sends them to Supabase Auth over HTTPS. Every supplier Excel preview and generate request requires a verified Supabase Auth session and a server-side `ALLOWED_EMAILS` match; public signup should remain disabled.

The earlier `ENABLE_REAL_QUOTATIONS` environment variable is **deprecated**. Authenticated processing becomes available when the three Supabase configuration variables are present and the account passes server authorization. Password `123456` is trivially guessable; replace it with a stronger individual password through the future user settings flow. Supabase rate limiting and login-abuse monitoring are recommended.

## Preview hasil quotation (v2.3.1)

Pengguna memilih hasil **CommVerge**, **Smartoptics**, atau **Keduanya** tanpa kode teknis format pada antarmuka. File sumber boleh menggunakan salah satu format; tipe file menentukan aturan perhitungan, bukan format output.

Preview hanya menampilkan **quotation hasil generator**, dalam bentuk cuplikan visual satu dokumen (logo yang tersedia, header, customer, item awal/akhir, total, dan terms). File sumber tetap dianalisis di backend, tetapi **tidak ditampilkan sebagai dokumen kedua**. Tujuh pemeriksaan SKU, kuantitas, deskripsi, harga, subtotal, perubahan margin, dan selisih muncul sebagai daftar angka 01–07 di bawah preview. Komparasi memeriksa semua item, termasuk item yang tidak terlihat dalam cuplikan.

Hasil download Excel, PDF, dan ZIP menggunakan nama CommVerge/Smartoptics dalam filename. Identifier internal tetap dipakai untuk kompatibilitas API dan generator, namun tidak ditampilkan dalam label pengguna atau nama file. Preview visual merupakan ringkasan dari sel XLSX output, bukan rasterisasi pixel-perfect workbook Excel.

## Production architecture

- **DEMO**: synthetic pricing and synthetic sample XLSX/ZIP; no supplier files or real quotation metadata.
- **REAL (requires configured Auth)**: Supabase Auth verified server-side, with explicit `ALLOWED_EMAILS` allowlist. Needs a **separate Supabase project**; no secrets in this repository.
- **Source XLSX**: authenticated upload, parsed in-memory. Smartoptics Detailed Quote original is repriced from supplier cost with gross margin (not markup). CommVerge T01 already priced is repaired using original existing unit prices without extra margin. QA, anti-double-margin checks and DRAFT review always apply.
- **T02 output**: Smartoptics input keeps its original workbook styling; CommVerge input uses a sanitized synthetic Smartoptics-style worksheet layout, replacing all demo rows and prices, clearly stamped as a cross-formatted DRAFT, not supplier-issued. The original Smartoptics vendor cost/list price is unknown in CommVerge-priced files and is intentionally left blank.
- **PDF (A4)**: available from authenticated quotation screen, using the exact same validated source, margin, and FX as Excel. A built-in, authenticated PDF preview renders prior to download. The PDF uses a compact A4 quotation layout; it is not a pixel-for-pixel print of the original XLSX template. T01 includes the official CommVerge logo supplied via the private master. T02 is clearly labelled as a repriced DRAFT, not supplier-issued. `BOTH` downloads a ZIP of two Excel files or two PDF files, according to the selected output type.
- **Output selection**: Excel (`.xlsx`) or PDF (`.pdf`), with ZIP packaging for BOTH. All real operations require Supabase Auth and server allowlist; supplier workbooks are processed in-memory only.

- **T01 output**: reads a reusable approved CommVerge master from Supabase **private Storage** bucket `qms-templates`, object `T01_MASTER_CLEAN.xlsx`, using the verified user's JWT. The master is never placed in the public repository. Only the two approved users can read that exact object; uploads/changes require administrator access through Supabase Dashboard.
- **T01 onboarding (one time)**: Admin opens Supabase Storage for project `endttqjueyeyropzdosy`, selects private bucket `qms-templates`, uploads the vetted master as `T01_MASTER_CLEAN.xlsx`. The master was recovered and verified from the earlier private portal package; don't upload a completed customer quotation instead. Once present, the user sees `Master CommVerge T01 sudah tersedia` after login and does not need to upload a second file.
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

- Source adapters: (1) original Smartoptics Detailed Quote (gross margin is calculated from original supplier unit costs), and (2) previously priced single-sheet CommVerge T01 customer quotation (existing unit selling prices are preserved, subtotal formula errors are recalculated from quantity × unit price, gross margin cannot be recalculated without supplier costs). Both source types can be exported as T01, T02 or BOTH. Pre-priced CommVerge sources preserve existing selling prices across target formats, with no second margin or invented supplier cost. A sanitized synthetic T02 layout is used for cross-format preview/export; this is not a verified pixel-perfect Smartoptics master. Smartoptics-supplier → CommVerge branded T01 requires a private approved CommVerge master in storage.
- T01 master must be provided separately and rights to use its brand/template must be confirmed by the issuer.
- Demo is **not** equivalent to successful authenticated end-to-end verification; real supplier XLSX output still requires live QA.
- No built-in pricing/tax assumptions beyond explicitly entered margin and FX.
