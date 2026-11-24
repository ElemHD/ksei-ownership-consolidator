#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolidates monthly KSEI pipe-delimited securities ownership snapshots (Balancepos*.txt)
into a multi-sheet, Stata-ready Excel workbook.
"""

import os
import sys
import re
import glob
import time
import argparse
from pathlib import Path
from typing import List, Tuple, Optional

import pandas as pd
import numpy as np


# KSEI raw files contain two identically named 'Total' columns (index 14 for local,
# index 24 for foreign); assign distinct raw names before mapping to Stata snake_case.
RAW_COLUMN_NAMES = [
    'Date', 'Code', 'Type', 'Sec. Num', 'Price',
    'Local IS', 'Local CP', 'Local PF', 'Local IB', 'Local ID',
    'Local MF', 'Local SC', 'Local FD', 'Local OT', 'Total_Local_Raw',
    'Foreign IS', 'Foreign CP', 'Foreign PF', 'Foreign IB', 'Foreign ID',
    'Foreign MF', 'Foreign SC', 'Foreign FD', 'Foreign OT', 'Total_Foreign_Raw'
]

COLUMN_RENAME_MAP = {
    'Date': 'report_date_str',
    'Code': 'ticker',
    'Type': 'sec_type',
    'Sec. Num': 'sec_num_listed',
    'Price': 'price_ksei',
    'Local IS': 'local_is',
    'Local CP': 'local_cp',
    'Local PF': 'local_pf',
    'Local IB': 'local_ib',
    'Local ID': 'local_id',
    'Local MF': 'local_mf',
    'Local SC': 'local_sc',
    'Local FD': 'local_fd',
    'Local OT': 'local_ot',
    'Total_Local_Raw': 'local_total',
    'Foreign IS': 'foreign_is',
    'Foreign CP': 'foreign_cp',
    'Foreign PF': 'foreign_pf',
    'Foreign IB': 'foreign_ib',
    'Foreign ID': 'foreign_id',
    'Foreign MF': 'foreign_mf',
    'Foreign SC': 'foreign_sc',
    'Foreign FD': 'foreign_fd',
    'Foreign OT': 'foreign_ot',
    'Total_Foreign_Raw': 'foreign_total',
}

NUMERIC_COLUMNS = [
    'sec_num_listed', 'price_ksei',
    'local_is', 'local_cp', 'local_pf', 'local_ib', 'local_id',
    'local_mf', 'local_sc', 'local_fd', 'local_ot', 'local_total',
    'foreign_is', 'foreign_cp', 'foreign_pf', 'foreign_ib', 'foreign_id',
    'foreign_mf', 'foreign_sc', 'foreign_fd', 'foreign_ot', 'foreign_total'
]

FINAL_COLUMN_ORDER = [
    'report_date_str', 'report_date', 'year_month', 'year', 'month',
    'ticker', 'sec_type', 'sec_num_listed', 'price_ksei',
    'local_is', 'local_cp', 'local_pf', 'local_ib', 'local_id',
    'local_mf', 'local_sc', 'local_fd', 'local_ot', 'local_total',
    'foreign_is', 'foreign_cp', 'foreign_pf', 'foreign_ib', 'foreign_id',
    'foreign_mf', 'foreign_sc', 'foreign_fd', 'foreign_ot', 'foreign_total',
    'source_file'
]

DEFAULT_SEARCH_PATHS = [
    os.path.join(".", "data"),
    os.path.join(".", "input"),
    os.path.join(".", "KepemilikanEfek_2020-2026"),
    os.path.join(".", "KepemilikanEfek_35bulan"),
    os.path.expanduser("~/Downloads/Dataset/KepemilikanEfek_2020-2026"),
    os.path.expanduser("~/Downloads/KepemilikanEfek_2020-2026"),
    os.path.expanduser("~/Downloads/KepemilikanEfek_35bulan"),
]


def find_default_input_dir() -> Optional[str]:
    env_dir = os.environ.get("KSEI_INPUT_DIR")
    if env_dir and os.path.isdir(env_dir):
        return env_dir
    for p in DEFAULT_SEARCH_PATHS:
        if os.path.exists(p) and os.path.isdir(p):
            return p
    return None


def collect_txt_files(root_dir: str) -> List[str]:
    path_obj = Path(root_dir)
    files = list(path_obj.rglob("*.txt"))
    valid_files = []
    for f in files:
        fname = f.name.lower()
        if "balancepos" in fname or fname.endswith(".txt"):
            valid_files.append(str(f))

    valid_files.sort(key=lambda x: os.path.basename(x))
    return valid_files


def parse_ksei_date(date_series: pd.Series) -> pd.Series:
    # KSEI standard date format is DD-MMM-YYYY (e.g. '31-OCT-2023').
    parsed = pd.to_datetime(date_series.astype(str).str.strip(), format='%d-%b-%Y', errors='coerce')
    if parsed.isna().any():
        fallback_mask = parsed.isna()
        parsed[fallback_mask] = pd.to_datetime(
            date_series[fallback_mask].astype(str).str.strip(),
            errors='coerce'
        )
    return parsed


def read_single_ksei_file(filepath: str) -> Optional[pd.DataFrame]:
    filename = os.path.basename(filepath)
    try:
        # Older KSEI releases occasionally use Latin-1 or CP1252 instead of UTF-8.
        encodings_to_try = ['utf-8', 'latin1', 'cp1252']
        raw_df = None
        for enc in encodings_to_try:
            try:
                raw_df = pd.read_csv(
                    filepath,
                    sep='|',
                    header=0,
                    names=RAW_COLUMN_NAMES,
                    dtype=str,
                    encoding=enc,
                    skip_blank_lines=True,
                    on_bad_lines='skip'
                )
                break
            except (UnicodeDecodeError, Exception):
                continue

        if raw_df is None or raw_df.empty:
            print(f"[PERINGATAN] File kosong atau gagal dibaca: {filename}")
            return None

        for col in raw_df.columns:
            raw_df[col] = raw_df[col].astype(str).str.strip()

        df = raw_df.rename(columns=COLUMN_RENAME_MAP)

        parsed_dates = parse_ksei_date(df['report_date_str'])
        df['report_date'] = parsed_dates.dt.strftime('%Y-%m-%d')
        df['year_month'] = parsed_dates.dt.strftime('%Y-%m')
        df['year'] = parsed_dates.dt.year.fillna(0).astype(int)
        df['month'] = parsed_dates.dt.month.fillna(0).astype(int)

        df['source_file'] = filename

        for col in NUMERIC_COLUMNS:
            if col in df.columns:
                cleaned = df[col].str.replace(',', '', regex=False).str.strip()
                df[col] = pd.to_numeric(cleaned, errors='coerce')

        return df

    except Exception as e:
        print(f"[ERROR] Terjadi kesalahan saat membaca file {filename}: {e}")
        return None


def process_ksei_data(
    input_dir: str,
    output_excel_path: str
) -> None:
    start_time = time.time()
    print("=" * 80)
    print("  KSEI SECURITIES OWNERSHIP CONSOLIDATOR (STATA-READY)")
    print("=" * 80)
    print(f"Direktori Input  : {os.path.abspath(input_dir)}")
    print(f"File Output Excel: {os.path.abspath(output_excel_path)}")
    print("-" * 80)

    txt_files = collect_txt_files(input_dir)
    if not txt_files:
        print(f"[FATAL ERROR] Tidak ditemukan file .txt di dalam: {input_dir}")
        sys.exit(1)

    print(f"Ditemukan {len(txt_files)} file data KSEI (.txt). Mulai membaca...")

    dfs = []
    for idx, f in enumerate(txt_files, 1):
        df_single = read_single_ksei_file(f)
        if df_single is not None and not df_single.empty:
            dfs.append(df_single)
            ym = df_single['year_month'].iloc[0] if 'year_month' in df_single else "N/A"
            print(f"  [{idx:02d}/{len(txt_files):02d}] {os.path.basename(f)} | Periode: {ym} | Baris: {len(df_single):,}")

    if not dfs:
        print("[FATAL ERROR] Tidak ada data yang berhasil dibaca.")
        sys.exit(1)

    full_df = pd.concat(dfs, ignore_index=True)
    total_raw_rows = len(full_df)
    print("-" * 80)
    print(f"Total baris seluruh instrumen terbaca: {total_raw_rows:,} baris.")

    available_cols = [c for c in FINAL_COLUMN_ORDER if c in full_df.columns]
    full_df = full_df[available_cols]

    # IDX common stocks use a strict 4-letter alphabetic ticker with sec_type == 'EQUITY'.
    # Warrants, rights, structured warrants, debt, and preferred shares are routed to a separate sheet.
    is_equity_type = full_df['sec_type'].str.upper() == 'EQUITY'
    is_pure_4letter = full_df['ticker'].str.match(r'^[A-Z]{4}$')

    pure_equity_mask = is_equity_type & is_pure_4letter
    equity_df = full_df[pure_equity_mask].copy()
    non_equity_df = full_df[~pure_equity_mask].copy()

    print(f"  - Saham Biasa Murni (EQUITY 4-Huruf): {len(equity_df):,} baris")
    print(f"  - Non-Equity & Derivatif/Lainnya    : {len(non_equity_df):,} baris")

    # Guard against overlapping monthly snapshots producing duplicate (ticker, year_month) panel keys.
    dup_mask = equity_df.duplicated(subset=['ticker', 'year_month'], keep='first')
    num_dups = dup_mask.sum()
    if num_dups > 0:
        print(f"[PERINGATAN] Terdeteksi {num_dups} observasi duplikat (ticker, year_month). Menghapus duplikasi...")
        equity_df = equity_df[~dup_mask].copy()
    else:
        print("  - Verifikasi Duplikasi: 0 duplikat ditemukan (Data konsisten & bersih).")

    equity_df = equity_df.sort_values(by=['ticker', 'report_date']).reset_index(drop=True)
    non_equity_df = non_equity_df.sort_values(by=['sec_type', 'ticker', 'report_date']).reset_index(drop=True)

    summary_list = []
    grouped = full_df.groupby(['year_month', 'sec_type'], as_index=False)
    for (ym, stype), group in grouped:
        summary_list.append({
            'year_month': ym,
            'sec_type': stype,
            'total_instruments': group['ticker'].nunique(),
            'total_rows': len(group),
            'sum_listed_shares': group['sec_num_listed'].sum(min_count=1),
            'sum_local_shares': group['local_total'].sum(min_count=1),
            'sum_foreign_shares': group['foreign_total'].sum(min_count=1),
            'sum_total_scripless': (group['local_total'] + group['foreign_total']).sum(min_count=1),
        })
    summary_df = pd.DataFrame(summary_list)
    summary_df = summary_df.sort_values(by=['year_month', 'sec_type']).reset_index(drop=True)

    print("-" * 80)
    print("Mengekspor seluruh tabel ke Excel Master...")

    output_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    excel_engine = 'xlsxwriter'
    try:
        import xlsxwriter
    except ImportError:
        excel_engine = 'openpyxl'

    print(f"Menggunakan Excel Engine: '{excel_engine}'")

    with pd.ExcelWriter(output_excel_path, engine=excel_engine) as writer:
        print(f"  [1/4] Menulis sheet 'ALL_EQUITY_STATA' ({len(equity_df):,} baris)...")
        equity_df.to_excel(writer, sheet_name='ALL_EQUITY_STATA', index=False)

        unique_years = sorted(equity_df['year'].unique())
        print(f"  [2/4] Menulis sheet tahunan untuk tahun: {unique_years}...")
        for yr in unique_years:
            yr_df = equity_df[equity_df['year'] == yr]
            sheet_title = f"EQUITY_{yr}"
            yr_df.to_excel(writer, sheet_name=sheet_title, index=False)
            print(f"        -> Sheet '{sheet_title}': {len(yr_df):,} baris")

        print(f"  [3/4] Menulis sheet 'NON_EQUITY_OR_DERIVATIVES' ({len(non_equity_df):,} baris)...")
        non_equity_df.to_excel(writer, sheet_name='NON_EQUITY_OR_DERIVATIVES', index=False)

        print(f"  [4/4] Menulis sheet 'SUMMARY_BY_MONTH_TYPE' ({len(summary_df):,} baris)...")
        summary_df.to_excel(writer, sheet_name='SUMMARY_BY_MONTH_TYPE', index=False)

    elapsed = time.time() - start_time
    file_size_mb = os.path.getsize(output_excel_path) / (1024 * 1024)

    print("=" * 80)
    print("SELESAI! KONSOLIDASI DATA KSEI BERHASIL DILAKUKAN")
    print("=" * 80)
    print(f"Lokasi File Master: {os.path.abspath(output_excel_path)}")
    print(f"Ukuran File Excel : {file_size_mb:.2f} MB")
    print(f"Waktu Pemrosesan  : {elapsed:.2f} detik")
    print("-" * 80)
    print("PANDUAN LANGKAH LANJUTAN DI STATA:")
    print("-" * 80)
    stata_code = f"""
* 1. Import sheet utama dari file Excel Master:
import excel "{os.path.abspath(output_excel_path)}", sheet("ALL_EQUITY_STATA") firstrow clear

* 2. Format tanggal dan deklarasi panel data:
gen date_stata = date(report_date, "YMD")
format date_stata %td
gen ym_stata = mofd(date_stata)
format ym_stata %tm
encode ticker, gen(id_emiten)
xtset id_emiten ym_stata

* 3. Hitung total saham scripless di KSEI (Lokal + Asing):
gen double total_scripless = local_total + foreign_total

* 4. Hitung proporsi kepemilikan ritel lokal (Local ID):
*    a. Proporsi terhadap total saham scripless (dalam sistem KSEI)
gen double prop_retail_scripless = local_id / total_scripless if total_scripless > 0

*    b. Proporsi terhadap total saham tercatat (Listed Shares)
gen double prop_retail_listed    = local_id / sec_num_listed  if sec_num_listed > 0
"""
    print(stata_code)
    print("=" * 80)


def main():
    parser = argparse.ArgumentParser(
        description="Konsolidator Data Kepemilikan Efek KSEI (.txt ke Master Excel Stata-Ready)"
    )
    parser.add_argument(
        '-i', '--input-dir',
        dest='input_dir',
        type=str,
        default=None,
        help="Path folder yang memuat kumpulan file Balancepos*.txt KSEI."
    )
    parser.add_argument(
        '-o', '--output',
        dest='output',
        type=str,
        default='Master_Raw_KSEI_Ownership.xlsx',
        help="Path file Excel output (.xlsx). Default: Master_Raw_KSEI_Ownership.xlsx"
    )

    args = parser.parse_args()

    input_directory = args.input_dir
    if not input_directory:
        detected = find_default_input_dir()
        if detected:
            input_directory = detected
            print(f"[INFO] Folder input otomatis terdeteksi: {input_directory}")
        else:
            print("[ERROR] Path folder input tidak ditentukan dan tidak ditemukan di lokasi default.")
            print(f"Lokasi default yang dicari: {DEFAULT_SEARCH_PATHS}")
            sys.exit(1)

    if not os.path.exists(input_directory):
        print(f"[ERROR] Direktori tidak ditemukan: {input_directory}")
        sys.exit(1)

    process_ksei_data(input_directory, args.output)


if __name__ == "__main__":
    main()
