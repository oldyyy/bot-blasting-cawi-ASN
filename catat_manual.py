"""
Cek & catat assignment yang diproses MANUAL, supaya bot tidak memprosesnya lagi.

  python catat_manual.py "NAMA KELUARGA"            # cek saja: sudah SUCCESS oleh bot atau belum?
  python catat_manual.py "NAMA KELUARGA" --catat    # catat sebagai SUCCESS (manual) ke logs/log_execute_manual.csv

Catatan: bot membaca log hanya saat mulai. Bot yang SEDANG berjalan tidak tahu catatan baru,
jadi cek dulu sebelum mengerjakan manual (dan sebaiknya kerjakan manual setelah bot selesai/berhenti).
"""
import argparse
import csv
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
LOG_DIR = BASE / "logs"
MANUAL_LOG = LOG_DIR / "log_execute_manual.csv"

ap = argparse.ArgumentParser()
ap.add_argument("nama", nargs="+", help="Nama Keluarga persis seperti di daftar_target.csv")
ap.add_argument("--catat", action="store_true", help="catat sebagai SUCCESS (manual)")
args = ap.parse_args()

targets = pd.read_csv(BASE / "daftar_target.csv", dtype=str)
logs = [pd.read_csv(f, dtype=str).assign(sumber=f.name) for f in LOG_DIR.glob("log_execute*.csv")]
log = pd.concat(logs, ignore_index=True)

for nama in args.nama:
    m = targets[targets["nama_keluarga"].str.upper() == nama.strip().upper()]
    if m.empty:
        print(f"[?] '{nama}' tidak ada di daftar_target.csv")
        continue
    t = m.iloc[0]
    hist = log[log["assignment_id"] == t["assignment_id"]].sort_values("waktu")
    sukses = hist[hist["status"] == "SUCCESS"]
    print(f"\n{t['nama_keluarga']} (no {t['no']}, {t['email']}, {t['telepon']})")
    for _, h in hist.iterrows():
        print(f"   {h['waktu']}  {h['status']:<22} [{h['sumber']}]")
    if not sukses.empty:
        print(f"   -> SUDAH SUCCESS pada {sukses['waktu'].iloc[0]}. JANGAN diproses manual (pesan akan terkirim dobel).")
        continue
    if not args.catat:
        print("   -> belum SUCCESS. Boleh diproses manual, lalu jalankan lagi dengan --catat.")
        continue
    new = not MANUAL_LOG.exists()
    with open(MANUAL_LOG, "a", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["waktu", "no", "assignment_id", "nama_keluarga", "email", "telepon", "status", "pesan"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), t["no"], t["assignment_id"],
                    t["nama_keluarga"], t["email"], t["telepon"], "SUCCESS", "diproses manual"])
    print("   -> dicatat sebagai SUCCESS (manual).")
