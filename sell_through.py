"""
Sell Through (distributor / RSP) -> Distributor_Sales.parquet, built in the cloud.

The ten regional workbooks ("1. Sumbagut Sell Through.xlsx" ...) are edited by hand on
the laptop. The laptop's mirror daemon uploads each one to
    gs://bucket_som/sales_parquet/raw/sell_through/Sell_Through_Offline_GT/
as soon as it is saved (sync.py in this repo does the same from OneDrive). This script is the cloud copy of the laptop's 1_raw_compiler.py:

    compile  every visible sheet of every workbook -> one table (same cleaning as the laptop)
    draft    sales_parquet/raw/sell_through/cloud/Distributor_Sales.parquet
    publish  sales_parquet/Distributor_Sales.parquet (the file BigQuery's
             raw_offline_distributor_sales reads), after validating it against the
             current production file and backing that up first.

    python sell_through.py run                 # gate -> compile -> draft -> validate -> publish
    python sell_through.py run --force         # skip the "anything new?" gate
    python sell_through.py run --no-publish    # stop after the draft
    python sell_through.py compile --folder DIR --out FILE   # local test, no GCS
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
import tempfile
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

WIB = ZoneInfo("Asia/Jakarta")
BUCKET = "bucket_som"
SRC_PREFIX = "sales_parquet/raw/sell_through/Sell_Through_Offline_GT/"
DRAFT = "sales_parquet/raw/sell_through/cloud/Distributor_Sales.parquet"
PRODUCTION = "sales_parquet/Distributor_Sales.parquet"
BACKUP_PREFIX = "sales_parquet/backup_small/Distributor_Sales-"   # one small file per publish
KEEP_BACKUPS = 30   # runs every 30 min now, so keep a few days of history

# Reference sheets that are not sales data. They are kept hidden in the workbooks, but are skipped by name as
# well, so unhiding one while working does not put its rows (and extra columns) in Distributor_Sales.parquet.
EXCLUDED_SHEETS = {"master_data"}

MIN_ROWS_VS_PRODUCTION = 0.90   # refuse a publish that would shrink the table by more than 10%
MIN_DATE_FILLED = 0.95          # share of rows with a parseable Date (~1% are empty Bali Nusra rows with Total 0, as on the laptop)


# --------------------------------------------------------------------------- #
# Compile - the laptop's 1_raw_compiler.py, one workbook load per file
# --------------------------------------------------------------------------- #

def compile_folder(folder: Path) -> pd.DataFrame:
    from openpyxl import load_workbook

    files = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"No .xlsx files in {folder}")

    frames = []
    for path in files:
        wb = load_workbook(path, read_only=True, data_only=False)
        visible = [ws.title for ws in wb.worksheets
                   if ws.sheet_state == "visible" and ws.title.strip().lower() not in EXCLUDED_SHEETS]
        skipped = [ws.title for ws in wb.worksheets
                   if ws.sheet_state == "visible" and ws.title.strip().lower() in EXCLUDED_SHEETS]
        if skipped:
            print(f"  skipping reference sheet(s) {skipped} in {path.name}")
        wb.close()
        if not visible:
            print(f"  no visible sheets in {path.name}, skipping")
            continue
        xl = pd.ExcelFile(path)
        for sheet in visible:
            df = xl.parse(sheet, header=0)
            # Headers are hand-maintained per sheet, so the same column can arrive
            # padded ('KLASIFIKASI OUTLET ') or not. Normalise before the concat.
            df.columns = [str(c).strip() for c in df.columns]
            df["SourceFile"] = path.name
            df["SheetName"] = sheet
            frames.append(df)
        print(f"  {path.name}: {len(visible)} sheet(s)")
    if not frames:
        raise SystemExit("No data from any visible sheet")

    out = pd.concat(frames, ignore_index=True)
    unnamed = [c for c in out.columns if c.startswith("Unnamed:")]
    if unnamed:
        print(f"  dropping unnamed columns: {unnamed}")
        out = out.drop(columns=unnamed)

    # Clean string columns (newlines out, trimmed) - every column goes through astype(str),
    # exactly as the laptop does, so missing cells become the text 'nan'.
    for col in out.columns:
        out[col] = out[col].astype(str).str.replace(r"[\n\r]+", " ", regex=True).str.strip()

    total_col = next((c for c in out.columns if c.lower() in ("total", "quantity")), None)
    if total_col:
        out[total_col] = pd.to_numeric(out[total_col], errors="coerce").fillna(0).astype(float)
    else:
        print("  WARNING: no 'Total' or 'Quantity' column found")

    if "Date" in out.columns:
        out["Date"] = pd.to_datetime(out["Date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    return out


# --------------------------------------------------------------------------- #
# GCS steps
# --------------------------------------------------------------------------- #

def _bucket():
    from google.cloud import storage
    return storage.Client().bucket(BUCKET)


def source_blobs(bucket):
    return [b for b in bucket.list_blobs(prefix=SRC_PREFIX)
            if b.name.lower().endswith(".xlsx") and not b.name.rsplit("/", 1)[-1].startswith("~$")]


def gate(bucket, force: bool) -> tuple[bool, str]:
    srcs = source_blobs(bucket)
    if not srcs:
        return False, f"no workbooks under gs://{BUCKET}/{SRC_PREFIX}"
    newest = max(b.updated for b in srcs)
    draft = bucket.get_blob(DRAFT)
    print(f"{len(srcs)} workbook(s), newest uploaded {newest.astimezone(WIB):%Y-%m-%d %H:%M} WIB")
    print(f"draft: {draft.updated.astimezone(WIB):%Y-%m-%d %H:%M} WIB" if draft else "draft: (none yet)")
    if force:
        return True, "forced"
    if draft is None:
        return True, "no draft yet"
    if newest > draft.updated:
        return True, "a workbook is newer than the draft"
    return False, "nothing newer than the draft"


def download_sources(bucket, dest: Path) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    for b in source_blobs(bucket):
        b.download_to_filename(str(dest / b.name.rsplit("/", 1)[-1]), timeout=900)


def validate(new: pd.DataFrame, bucket) -> list[str]:
    """Reasons NOT to publish; empty = fine."""
    import io
    problems = []
    prod = bucket.get_blob(PRODUCTION)
    if prod is not None:
        import pyarrow.parquet as pq
        with prod.open("rb") as fh:
            meta = pq.ParquetFile(fh).metadata
            prod_cols, prod_rows = meta.schema.to_arrow_schema().names, meta.num_rows
        if list(new.columns) != list(prod_cols):
            problems.append(f"columns differ from production: new={list(new.columns)} production={prod_cols}")
        if len(new) < MIN_ROWS_VS_PRODUCTION * prod_rows:
            problems.append(f"{len(new):,} rows is more than 10% below production's {prod_rows:,}")
    filled = (new["Date"] != "").mean() if "Date" in new.columns and len(new) else 0
    if filled < MIN_DATE_FILLED:
        problems.append(f"only {filled:.1%} of rows have a valid Date")
    return problems


def publish(draft_blob_name: str, bucket) -> None:
    stamp = dt.datetime.now(WIB).strftime("%Y%m%d-%H%M")
    prod = bucket.get_blob(PRODUCTION)
    if prod is not None:
        bucket.copy_blob(prod, bucket, f"{BACKUP_PREFIX}{stamp}.parquet")
        print(f"backed up production -> {BACKUP_PREFIX}{stamp}.parquet")
    bucket.copy_blob(bucket.blob(draft_blob_name), bucket, PRODUCTION)
    print(f"published {draft_blob_name} -> {PRODUCTION}")
    backups = sorted(bucket.list_blobs(prefix=BACKUP_PREFIX), key=lambda b: b.name)
    for old in backups[:-KEEP_BACKUPS]:
        old.delete()
        print(f"pruned {old.name}")


def unchanged(new: pd.DataFrame, bucket) -> bool:
    """True when production already holds exactly this table (no pointless backup + rewrite)."""
    prod = bucket.get_blob(PRODUCTION)
    if prod is None:
        return False
    try:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "production.parquet"
            prod.download_to_filename(str(path), timeout=900)
            old = pd.read_parquet(path)
        return (list(old.columns) == list(new.columns) and len(old) == len(new)
                and old.astype(object).equals(new.astype(object)))
    except Exception as e:  # any doubt -> publish
        print(f"  could not compare with production ({e}); publishing")
        return False


def run(args) -> int:
    bucket = _bucket()
    go, reason = gate(bucket, args.force)
    print(f"build: {go} ({reason})")
    if not go:
        return 0
    with tempfile.TemporaryDirectory() as tmp:
        inputs, out = Path(tmp) / "inputs", Path(tmp) / "Distributor_Sales.parquet"
        download_sources(bucket, inputs)
        df = compile_folder(inputs)
        print(f"compiled {len(df):,} rows, {df['Date'].replace('', pd.NA).dropna().min()} .. "
              f"{df['Date'].replace('', pd.NA).dropna().max()}")
        df.to_parquet(out, index=False, engine="pyarrow")
        bucket.blob(DRAFT).upload_from_filename(str(out), timeout=900)
        print(f"draft uploaded -> {DRAFT} ({out.stat().st_size / 1e6:.1f} MB)")
        if args.no_publish:
            return 0
        problems = validate(df, bucket)
        if problems:
            print("NOT published:")
            for p in problems:
                print("  -", p)
            return 1
        if unchanged(df, bucket):
            print("compiled table is identical to production - nothing to publish")
            return 0
        publish(DRAFT, bucket)
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run")
    r.add_argument("--force", action="store_true")
    r.add_argument("--no-publish", action="store_true")
    c = sub.add_parser("compile")
    c.add_argument("--folder", type=Path, required=True)
    c.add_argument("--out", type=Path, required=True)
    a = ap.parse_args(argv)
    if a.cmd == "compile":
        df = compile_folder(a.folder)
        df.to_parquet(a.out, index=False, engine="pyarrow")
        print(f"{len(df):,} rows -> {a.out}")
        return 0
    return run(a)


if __name__ == "__main__":
    sys.exit(main())
