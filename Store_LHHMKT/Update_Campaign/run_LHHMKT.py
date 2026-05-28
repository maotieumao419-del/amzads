"""
run_LHHMKT.py
================
MASTER RUNNER – Amazon Ads Campaign Rename & Upload Pipeline | Store: LHHMKT

Prefix tên Campaign: SKU trực tiếp (ví dụ: MY_SKU_123_SP02_AUTO_Nguyen)

Chạy toàn bộ 4 bước pipeline chỉ bằng một lệnh duy nhất:

    BƯỚC 0 – Load & làm sạch BulkSheetExport_*.xlsx
    BƯỚC 1 – Phân tách SKU & phát hiện chiến dịch phức tạp      (module1_isolation)
    BƯỚC 2+3 – Đặt tên theo SOP & tạo file update per-SKU       (module2_naming + module3_upload_generator)
    BƯỚC 4 – Gộp tất cả file per-SKU thành 1 Master Upload file (merge_upload_files)

Cây thư mục output:
    data/
    ├── input/
    │   └── BulkSheetExport_*.xlsx           ← ĐẶT FILE AMAZON VÀO ĐÂY
    ├── sku_files/
    │   ├── SKU_<sku>.xlsx                   ← Bước 1: toàn bộ row của từng SKU
    │   └── Complex_Campaigns_Review.xlsx    ← Bước 1: chiến dịch cần review tay
    ├── upload_files/
    │   └── Update_CampaignName_<sku>.xlsx   ← Bước 3: file update từng SKU
    └── master/
        ├── Master_Amazon_Bulksheet_Update_Upload.xlsx   ← Bước 4: file upload cuối
        └── Master_Amazon_Bulksheet_Update_Upload.csv    ← Bước 4: bản CSV (utf-8-sig)

Cách chạy:
    python run_LHHMKT.py                              # tự động tìm file
    python run_LHHMKT.py --bulk data/input/file.xlsx
    python run_LHHMKT.py --template path/to/tpl.xlsx
    python run_LHHMKT.py --no-csv                    # bỏ qua output CSV
    python run_LHHMKT.py --skip-merge                # chỉ chạy Bước 1-3

Args (tất cả đều có thể bỏ qua – pipeline sẽ tự tìm):
    --bulk          Đường dẫn đến file BulkSheetExport_*.xlsx
    --template      Đường dẫn đến Amazon template .xlsx (fallback về plain writer)
    --no-csv        Không tạo file .csv ở Bước 4
    --skip-merge    Bỏ qua Bước 4 (chỉ tạo file per-SKU)
"""

import argparse
import glob
import os
import sys
from datetime import datetime

# ── Import tất cả các module nội bộ ─────────────────────────────────────────
from utils                    import find_latest_bulksheet, load_bulksheet, safe_filename
from module1_isolation        import isolate_campaigns
from module2_naming           import generate_rename_map
from module3_upload_generator import generate_upload_file
from merge_upload_files       import merge_upload_files


# ── Cấu trúc thư mục ─────────────────────────────────────────────────────────
BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
DATA_DIR         = os.path.join(BASE_DIR, "data")
SKU_FILES_DIR    = os.path.join(DATA_DIR, "sku_files")     # output Bước 1
UPLOAD_FILES_DIR = os.path.join(DATA_DIR, "upload_files")  # output Bước 3
MASTER_DIR       = os.path.join(DATA_DIR, "master")        # output Bước 4

# Đường dẫn mặc định tới Amazon template (2 cấp trên thư mục hiện tại)
_DEFAULT_TEMPLATE = os.path.join(
    BASE_DIR, "..", "RULE&TEMPLATE",
    "AmazonAdvertisingBulksheetSellerTemplate (3).xlsx",
)


# ── Phân tích tham số CLI ─────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Amazon Ads – Master Pipeline Runner (All-in-One)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument(
        "--bulk", metavar="FILE",
        help="Đường dẫn file BulkSheetExport_*.xlsx. Mặc định: tự tìm trong data/input/",
    )
    p.add_argument(
        "--template", metavar="FILE",
        help="Đường dẫn Amazon template .xlsx. Mặc định: tự tìm hoặc dùng plain writer.",
    )
    p.add_argument(
        "--no-csv", action="store_true", default=False,
        help="Bỏ qua việc tạo file .csv ở Bước 4.",
    )
    p.add_argument(
        "--skip-merge", action="store_true", default=False,
        help="Bỏ qua Bước 4 (không tạo Master Upload file).",
    )
    return p.parse_args()


# ── Tìm Amazon template ───────────────────────────────────────────────────────

def _find_template(explicit: str | None) -> str | None:
    """Tìm file Amazon seller template theo thứ tự ưu tiên."""
    if explicit:
        path = os.path.abspath(explicit)
        if os.path.exists(path):
            return path
        print(f"[PIPELINE] CẢNH BÁO – Template không tìm thấy: {explicit}")
        return None

    # Thử đường dẫn mặc định
    default = os.path.normpath(_DEFAULT_TEMPLATE)
    if os.path.exists(default):
        print(f"[PIPELINE] Template tìm thấy: {default}")
        return default

    # Tìm rộng hơn 2 cấp trên
    for hit in glob.glob(
        os.path.join(BASE_DIR, "..", "**", "AmazonAdvertising*.xlsx"),
        recursive=True,
    ):
        print(f"[PIPELINE] Template tìm thấy (auto-search): {hit}")
        return os.path.abspath(hit)

    print("[PIPELINE] CẢNH BÁO – Không tìm thấy Amazon template → dùng plain writer.")
    return None


# ── Hàm in tiêu đề phần ──────────────────────────────────────────────────────

def _section(title: str) -> None:
    print()
    print("┌" + "─" * 68 + "┐")
    print(f"│  {title:<66}│")
    print("└" + "─" * 68 + "┘")


# ── Orchestrator chính ────────────────────────────────────────────────────────

def run_pipeline(
    bulk_path:     str,
    template_path: str | None,
    write_csv:     bool = True,
    skip_merge:    bool = False,
) -> None:
    """
    Chạy toàn bộ pipeline 4 bước theo thứ tự nghiêm ngặt.

    Args:
        bulk_path:     Đường dẫn tuyệt đối đến file BulkSheetExport_*.xlsx.
        template_path: Đường dẫn Amazon template, hoặc None để dùng plain writer.
        write_csv:     Nếu True, Bước 4 cũng tạo thêm file .csv.
        skip_merge:    Nếu True, bỏ qua Bước 4.
    """
    start_time = datetime.now()

    # ── Banner ────────────────────────────────────────────────────────────────
    print()
    print("╔" + "═" * 68 + "╗")
    print("║     AMAZON ADS – LHHMKT PIPELINE RUNNER (SKU PREFIX)             ║")
    print("╠" + "═" * 68 + "╣")
    tpl_label = os.path.basename(template_path) if template_path else "(plain writer)"
    print(f"║  Bulk file  : {os.path.basename(bulk_path):<54}║")
    print(f"║  Template   : {tpl_label:<54}║")
    print(f"║  Bắt đầu   : {start_time.strftime('%Y-%m-%d %H:%M:%S'):<54}║")
    print("╚" + "═" * 68 + "╝")

    # ══════════════════════════════════════════════════════════════════════════
    # BƯỚC 0 – Đọc & làm sạch BulkSheet
    # ══════════════════════════════════════════════════════════════════════════
    _section("BƯỚC 0 – Đọc & Làm Sạch Dữ Liệu BulkSheet")
    df_raw = load_bulksheet(bulk_path)
    print(f"[BƯỚC 0] ✓ Đã tải {len(df_raw):,} dòng từ file BulkSheet.")

    # ══════════════════════════════════════════════════════════════════════════
    # BƯỚC 1 – Phân tách SKU & phát hiện chiến dịch phức tạp
    # ══════════════════════════════════════════════════════════════════════════
    _section("BƯỚC 1 – Phân Tách SKU & Phát Hiện Chiến Dịch Phức Tạp")

    standard_sku_map, complex_cids = isolate_campaigns(
        df=df_raw,
        output_dir=SKU_FILES_DIR,
    )

    if not standard_sku_map:
        print(
            "\n[BƯỚC 1] Không còn chiến dịch tiêu chuẩn nào để xử lý. "
            "Pipeline kết thúc sớm."
        )
        return

    print(
        f"\n[BƯỚC 1] ✓ {len(standard_sku_map)} SKU tiêu chuẩn | "
        f"{len(complex_cids)} chiến dịch phức tạp đã cách ly."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # BƯỚC 2 + 3 – Đặt tên SOP & Tạo file upload per-SKU
    # ══════════════════════════════════════════════════════════════════════════
    _section("BƯỚC 2+3 – Đặt Tên SOP & Tạo File Update Từng SKU")

    successful_skus:       list[str] = []
    failed_skus:           list[str] = []
    total_campaigns_named: int       = 0
    total_upload_files:    int       = 0

    for sku, campaign_ids in sorted(standard_sku_map.items()):
        print(f"\n  ▶ SKU: '{sku}'  ({len(campaign_ids)} chiến dịch)")

        # Đọc lại file per-SKU mà Bước 1 đã ghi ra
        sku_file = os.path.join(SKU_FILES_DIR, f"SKU_{safe_filename(sku)}.xlsx")

        if not os.path.exists(sku_file):
            print(f"    [LỖI] Không tìm thấy file: {sku_file} → BỎ QUA.")
            failed_skus.append(sku)
            continue

        try:
            df_sku = load_bulksheet(sku_file)
        except Exception as exc:
            print(f"    [LỖI] Không đọc được '{sku_file}': {exc} → BỎ QUA.")
            failed_skus.append(sku)
            continue

        # ── Bước 2: Tạo rename map ────────────────────────────────────────────
        rename_map = generate_rename_map(df_sku, sku)

        if not rename_map:
            print(f"    [CẢNH BÁO] Không tạo được tên cho SKU '{sku}' → BỎ QUA Bước 3.")
            failed_skus.append(sku)
            continue

        total_campaigns_named += len(rename_map)

        # ── Bước 3: Ghi file update ──────────────────────────────────────────
        out_path = generate_upload_file(
            df=df_sku,
            rename_map=rename_map,
            sku=sku,
            output_dir=UPLOAD_FILES_DIR,
            template_path=template_path,
        )

        if out_path:
            successful_skus.append(sku)
            total_upload_files += 1
        else:
            failed_skus.append(sku)

    print(
        f"\n[BƯỚC 2+3] ✓ {total_campaigns_named} chiến dịch đã đặt tên | "
        f"{total_upload_files} file update đã tạo."
    )

    # ══════════════════════════════════════════════════════════════════════════
    # BƯỚC 4 – Gộp tất cả per-SKU → Master Upload file
    # ══════════════════════════════════════════════════════════════════════════
    master_xlsx_path: str | None = None

    if skip_merge:
        _section("BƯỚC 4 – Gộp Master Upload (ĐÃ BỎ QUA theo --skip-merge)")
        print("[BƯỚC 4] Bỏ qua theo yêu cầu. Các file per-SKU vẫn được tạo ở Bước 3.")
    elif total_upload_files == 0:
        _section("BƯỚC 4 – Gộp Master Upload (BỎ QUA – không có file per-SKU)")
        print("[BƯỚC 4] Không có file per-SKU nào được tạo. Bỏ qua bước merge.")
    else:
        _section("BƯỚC 4 – Gộp Tất Cả File Per-SKU Thành Master Upload")
        master_xlsx_path = merge_upload_files(
            input_dir  = UPLOAD_FILES_DIR,
            output_dir = MASTER_DIR,
            write_csv  = write_csv,
            store_name = "LHHMKT",
            task_name  = "UpdateCampaigns",
        )
        if master_xlsx_path:
            print(f"[BƯỚC 4] ✓ Master Upload file đã sẵn sàng để upload lên Amazon.")

    # ══════════════════════════════════════════════════════════════════════════
    # TỔNG KẾT CUỐI
    # ══════════════════════════════════════════════════════════════════════════
    elapsed = (datetime.now() - start_time).total_seconds()

    print()
    print("╔" + "═" * 68 + "╗")
    print("║  PIPELINE HOÀN TẤT – TỔNG KẾT                                     ║")
    print("╠" + "═" * 68 + "╣")
    print(f"║  Tổng SKU tiêu chuẩn            : {len(standard_sku_map):<34}║")
    print(f"║  Chiến dịch phức tạp (cần review): {len(complex_cids):<34}║")
    print(f"║  Chiến dịch đã đặt tên (Bước 2) : {total_campaigns_named:<34}║")
    print(f"║  File per-SKU đã tạo (Bước 3)   : {total_upload_files:<34}║")
    if not skip_merge and master_xlsx_path:
        master_label = os.path.basename(master_xlsx_path)
        print(f"║  Master Upload file (Bước 4)    : {master_label:<34}║")
    if failed_skus:
        print(f"║  SKU lỗi / bỏ qua               : {len(failed_skus):<34}║")
    print(f"║  Thời gian chạy                 : {elapsed:.1f}s{'':<31}║")
    print("╠" + "═" * 68 + "╣")
    print(f"║  Bước 1 → {SKU_FILES_DIR:<58}║")
    print(f"║  Bước 3 → {UPLOAD_FILES_DIR:<58}║")
    if not skip_merge:
        print(f"║  Bước 4 → {MASTER_DIR:<58}║")
    print("╚" + "═" * 68 + "╝")

    # Cảnh báo quan trọng
    if complex_cids:
        complex_file = os.path.join(SKU_FILES_DIR, "Complex_Campaigns_Review.xlsx")
        print(f"\n⚠️  CẦN XEM THỦ CÔNG: {len(complex_cids)} chiến dịch phức tạp")
        print(f"   File: {complex_file}")

    if failed_skus:
        print(f"\n⚠️  {len(failed_skus)} SKU bị lỗi hoặc bỏ qua:")
        for s in failed_skus:
            print(f"   • {s}")

    print()


# ── Entry Point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    args = _parse_args()

    # ── Resolve đường dẫn BulkSheet ──────────────────────────────────────────
    if args.bulk:
        if not os.path.exists(args.bulk):
            print(f"[PIPELINE] LỖI – Không tìm thấy file: {args.bulk}")
            sys.exit(1)
        bulk_xlsx = os.path.abspath(args.bulk)
    else:
        try:
            bulk_xlsx = find_latest_bulksheet(BASE_DIR)
            print(f"[PIPELINE] Auto-discover: {bulk_xlsx}")
        except FileNotFoundError as exc:
            print(exc)
            sys.exit(1)

    # ── Resolve template ──────────────────────────────────────────────────────
    template_xlsx = _find_template(args.template)

    # ── Chạy pipeline ─────────────────────────────────────────────────────────
    run_pipeline(
        bulk_path     = bulk_xlsx,
        template_path = template_xlsx,
        write_csv     = not args.no_csv,
        skip_merge    = args.skip_merge,
    )
