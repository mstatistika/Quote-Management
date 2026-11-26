"""Read-only comparison of uploaded XLSX with the actual generated XLSX draft.

Only called behind Supabase authentication. No vendor workbook is persisted.
Expected price adjustments are distinguished from unexpected item/quantity changes.
"""
from __future__ import annotations
import io
import zipfile
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
    return {
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
