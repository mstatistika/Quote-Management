"""Read-only comparison of uploaded XLSX with the actual generated XLSX draft.

Only called behind Supabase authentication. No vendor workbook is persisted.
Expected price adjustments are distinguished from unexpected item/quantity changes.
"""
from __future__ import annotations
import io
import zipfile
import base64
import datetime as dt
from decimal import Decimal, InvalidOperation

from quote_engine import QuoteError, SLOTS, generate, source, workbook, sheet_and_shared, value


def _same_amount(a, b):
    try:
        return abs(Decimal(str(a)) - Decimal(str(b))) <= Decimal("0.02")
    except (ValueError, InvalidOperation, TypeError):
        return False


def _sheet_data(raw):
    zipped = workbook(raw)
    sheet, strings = sheet_and_shared(zipped)
    return sheet, strings



def _logo(z, is_commverge):
    """Return only a bounded raster logo from an already-authorized XLSX."""
    if not is_commverge:
        return None
    for filename in ("xl/media/image.jpg", "xl/media/image.jpeg", "xl/media/image1.png", "xl/media/image.png"):
        if filename not in z.namelist():
            continue
        item = z.getinfo(filename)
        if item.file_size > 180_000:
            continue
        data = z.read(filename)
        if data[:3] == b"\\xff\\xd8\\xff":
            mime = "image/jpeg"
        elif data.startswith(b"\\x89PNG\\r\\n\\x1a\\n"):
            mime = "image/png"
        else:
            continue
        return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"
    return None


def _date_value(value_text):
    value_text = str(value_text or "").strip()
    if not value_text:
        return "Belum diisi"
    try:
        number = float(value_text)
        if 30000 <= number <= 85000:
            return (dt.date(1899, 12, 30) + dt.timedelta(days=int(number))).isoformat()
    except (ValueError, OverflowError):
        pass
    return value_text[:42]


def _document_snapshot(raw, code, rows, original=False):
    """Visual single-page excerpt based on cells in the *actual* workbook."""
    zipped = workbook(raw)
    sheet, shared = sheet_and_shared(zipped)
    def get(addr):
        return value(sheet, shared, addr).strip()
    t01 = code == "T01"
    if t01:
        meta = {
            "company": get("D6") or "CommVerge Solutions",
            "customer": get("C10"),
            "attention": get("C11"),
            "project": get("F12"),
            "reference": get("P10") or get("G11"),
            "date": _date_value(get("I10")),
            "delivery": get("D67") or get("C13"),
            "payment": get("D68"),
            "validity": get("D70"),
            "remarks": get("B64"),
        }
    else:
        meta = {
            "company": "Smartoptics",
            "customer": get("E9"),
            "attention": get("E8"),
            "project": get("E11"),
            "reference": get("L14"),
            "date": _date_value(get("E92")),
            "delivery": get("E93"),
            "payment": get("E94"),
            "validity": get("E95"),
            "remarks": get("E97"),
        }
    count = len(rows)
    if count <= 7:
        excerpt = rows
        omitted = 0
    else:
        excerpt = rows[:4] + rows[-2:]
        omitted = count - 6
    # The source grand total may be an Excel error; never replace it silently.
    subtotal_addr = "H45" if t01 else "M89"
    original_total = get(subtotal_addr)
    return {
        "template": code,
        "meta": meta,
        "logo": _logo(zipped, t01),
        "item_count": count,
        "omitted": omitted,
        "first_count": min(count, 4) if omitted else count,
        "items": [{k: r.get(k, "") for k in ("row", "sku", "description", "qty", "unit", "total", "status")} for r in excerpt],
        "total": original_total,
        "original": original,
        "preview_only": True,
    }


def compare_workbooks(raw, target_format="T01", margin=15, fx=18300, details="", master=None):
    """Return source and generated draft row grids with independently read XLSX cells."""
    if target_format not in ("T01", "T02", "BOTH"):
        raise QuoteError("Format quotation tidak valid.")
    q = source(raw, margin, fx)
    name, output, mime = generate(raw, target_format, margin, fx, details, master)
    viewed_format = "T01" if target_format == "BOTH" else target_format
    if target_format == "BOTH":
        with zipfile.ZipFile(io.BytesIO(output)) as archive:
            filename = next((n for n in archive.namelist() if "_T01_" in n and n.endswith(".xlsx")), None)
            if filename is None:
                raise QuoteError("Draft T01 tidak ditemukan dalam paket export.")
            output = archive.read(filename)
    original_sheet, original_strings = _sheet_data(raw)
    draft_sheet, draft_strings = _sheet_data(output)

    prepriced = q["source_type"] == "commverge_prepriced"
    source_columns = {"sku": "C", "description": "D", "qty": "F", "unit": "G", "total": "H"} if prepriced else {
        "sku": "D", "description": "E", "qty": "L", "unit": "G", "total": "M"}
    draft_columns = {"sku": "C", "description": "D", "qty": "F", "unit": "G", "total": "H"} if viewed_format == "T01" else {
        "sku": "D", "description": "E", "qty": "L", "unit": "G", "total": "M"}

    src_rows, target_rows = [], []
    sku_matches = qty_matches = description_matches = unit_matches = subtotal_matches = 0
    unexpected = repaired = adjusted = 0

    def read_row(sheet, shared, columns, row_number):
        data = {"row": row_number}
        for key, column in columns.items():
            data[key] = value(sheet, shared, f"{column}{row_number}")
        return data

    for index, item in enumerate(q["items"]):
        origin = read_row(original_sheet, original_strings, source_columns, item["row"])
        draft_row = SLOTS[index] if viewed_format == "T01" else item["row"]
        target = read_row(draft_sheet, draft_strings, draft_columns, draft_row)
        sku_equal = origin["sku"].strip() == target["sku"].strip()
        desc_equal = origin["description"].strip() == target["description"].strip()
        qty_equal = _same_amount(origin["qty"], target["qty"])
        price_valid = _same_amount(target["unit"], item["sell"])
        sum_valid = _same_amount(target["total"], item["total"])
        orig_price_same = _same_amount(origin["unit"], target["unit"])
        orig_subtotal_same = _same_amount(origin["total"], target["total"])

        sku_matches += sku_equal
        qty_matches += qty_equal
        description_matches += desc_equal
        unit_matches += price_valid
        subtotal_matches += sum_valid

        if not all((sku_equal, desc_equal, qty_equal, price_valid, sum_valid)):
            status, note = "PERIKSA", "Ada perbedaan yang tidak diharapkan. Jangan gunakan draft."
            unexpected += 1
        elif prepriced and not orig_subtotal_same:
            status, note = "DIPERBAIKI", "Subtotal dihitung ulang: QTY × harga jual existing; harga unit tidak berubah."
            repaired += 1
        elif not prepriced and (not orig_price_same or not orig_subtotal_same):
            status, note = "MARGIN", "Perbedaan harga diharapkan: kalkulasi gross margin dari biaya supplier."
            adjusted += 1
        else:
            status, note = "COCOK", "SKU, deskripsi, kuantitas dan nilai terverifikasi."

        origin["status"] = status
        target["status"] = status
        target["note"] = note
        src_rows.append(origin)
        target_rows.append(target)

    if unexpected:
        raise QuoteError(f"Draft tidak sesuai dengan sumber pada {unexpected} baris; export diblokir.")

    source_total_address = "H45" if prepriced else "M89"
    draft_total_address = "H45" if viewed_format == "T01" else "M89"
    source_code = "T01" if prepriced else "T02"
    source_doc = _document_snapshot(raw, source_code, src_rows, original=True)
    draft_doc = _document_snapshot(output, viewed_format, target_rows)
    return {
        "source_document": source_doc,
        "draft_document": draft_doc,
        "source_type": q["source_type"],
        "source_label": "COMMVERGE T01 · FILE DIUPLOAD" if prepriced else "SMARTOPTICS · FILE SUPPLIER DIUPLOAD",
        "draft_label": f"{'COMMVERGE' if viewed_format == 'T01' else 'SMARTOPTICS'} {viewed_format} · DRAFT GENERATOR",
        "format_viewed": viewed_format,
        "format_requested": target_format,
        "source_columns": source_columns,
        "draft_columns": draft_columns,
        "source_rows": src_rows,
        "draft_rows": target_rows,
        "source_total": value(original_sheet, original_strings, source_total_address),
        "draft_total": value(draft_sheet, draft_strings, draft_total_address),
        "same_layout": prepriced if viewed_format == "T01" else not prepriced,
        "summary": {
            "rows": len(src_rows),
            "sku_matches": sku_matches,
            "qty_matches": qty_matches,
            "description_matches": description_matches,
            "price_validated": unit_matches,
            "total_validated": subtotal_matches,
            "subtotals_repaired": repaired,
            "margin_adjustments": adjusted,
            "unexpected_differences": unexpected
        },
        "note": (
            "Harga jual pada file CommVerge dipertahankan. Subtotal yang tidak cocok atau error telah dihitung ulang. "
            "Tidak ada penambahan margin kedua."
            if prepriced else
            "Harga pada file sumber adalah harga supplier; draft menampilkan harga jual setelah gross margin. "
            "SKU, deskripsi dan QTY harus tetap sama."
        ) + (
            " Preview menunjukkan draft T01; paket download mencakup T01 dan T02."
            if target_format == "BOTH" else ""
        ),
    }
