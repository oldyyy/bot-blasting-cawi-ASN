"""
Rekap hasil broadcast: gabungkan semua logs/log_execute*.csv, ambil status TERAKHIR per assignment,
lalu tulis rekap_broadcast.xlsx (sheet Ringkasan, Berhasil, Perlu Manual, Belum Diproses).

  python rekap.py
"""
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
DONE = {"SUCCESS", "SKIP_SUDAH_CAWI"}

targets = pd.read_csv(BASE / "daftar_target.csv", dtype=str)
logs = [pd.read_csv(f, dtype=str).assign(sumber=f.name) for f in sorted((BASE / "logs").glob("log_execute*.csv"))]
log = pd.concat(logs, ignore_index=True) if logs else pd.DataFrame(columns=["waktu", "assignment_id", "status"])
log["percobaan"] = log.groupby("assignment_id")["assignment_id"].transform("size")
last = log.sort_values("waktu").groupby("assignment_id").tail(1)

cols = ["no", "assignment_id", "nama_keluarga", "alamat", "asn_dipilih", "email", "telepon"]
df = targets[cols].merge(last[["assignment_id", "waktu", "status", "pesan", "percobaan", "sumber"]],
                         on="assignment_id", how="left")
df["status"] = df["status"].fillna("BELUM_DIPROSES")

berhasil = df[df["status"].isin(DONE)]
belum = df[df["status"] == "BELUM_DIPROSES"]
manual = df[~df["status"].isin(DONE | {"BELUM_DIPROSES"})]

ringkas = df["status"].value_counts().rename_axis("status").reset_index(name="jumlah")
out = BASE / "rekap_broadcast.xlsx"
with pd.ExcelWriter(out) as xw:
    ringkas.to_excel(xw, sheet_name="Ringkasan", index=False)
    manual.to_excel(xw, sheet_name="Perlu Manual", index=False)
    berhasil.to_excel(xw, sheet_name="Berhasil", index=False)
    belum.to_excel(xw, sheet_name="Belum Diproses", index=False)
print(ringkas.to_string(index=False))
print(f"\nBerhasil {len(berhasil)} | perlu dicek/manual {len(manual)} | belum diproses {len(belum)} -> {out.name}")
