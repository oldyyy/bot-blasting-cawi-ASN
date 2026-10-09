"""
Bot broadcast CAWI FASIH-SM: ganti mode assignment OPEN ke CAWI dan kirim tautan ke responden.
Lihat README.md untuk cara pakai. Default DRY-RUN; klik final hanya dengan --execute.
"""
import argparse
import csv
import json
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
CONFIG_PATH = BASE / "config.json"
if not CONFIG_PATH.exists():
    sys.exit("config.json belum ada. Salin config.example.json menjadi config.json lalu isi.")
CONFIG = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
SRC_XLSX = (BASE / CONFIG["xlsx_path"]).resolve()
SHEET = CONFIG.get("sheet", "Daftar Assignment OPEN")
DATA_URL = CONFIG["data_url"]
PROFILE_DIR = BASE / ".chrome-profile"
LOG_DIR = BASE / "logs"
SHOT_DIR = BASE / "screenshots"
RECON_DIR = BASE / "recon"

EXCLUDE_KETERANGAN = re.compile(r"PENSIUN|TIDAK\s*DI\s*TEMUKAN", re.I)
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$")
PHONE_RE = re.compile(r"^628\d{7,11}$")

# Jeda (detik) - meniru ritme kerja manual dan menjaga beban server tetap ringan
STEP_DELAY = (1.0, 2.5)        # antar klik/aksi dalam satu assignment
ITEM_DELAY = (5.0, 15.0)       # antar assignment
LONG_BREAK_EVERY = (20, 35)    # setelah sekian assignment ...
LONG_BREAK = (40.0, 120.0)     # ... istirahat lebih lama
HOVER_DELAY = (0.1, 0.5)       # jeda antara kursor tiba di elemen dan klik
CLICK_HOLD_MS = (50, 180)      # lama tombol mouse ditekan (ms)


# ----------------------------------------------------------------------------- data
def parse_multi(value):
    """'NAMA A: x; NAMA B: y' -> [('NAMA A','x'), ('NAMA B','y')]; nilai tunggal -> [(None,'x')]"""
    if not isinstance(value, str) or not value.strip():
        return []
    out = []
    for part in value.split(";"):
        part = part.strip()
        if not part:
            continue
        if ":" in part and "@" not in part.split(":", 1)[0]:
            name, val = part.split(":", 1)
            out.append((name.strip().upper(), val.strip()))
        else:
            out.append((None, part))
    return out


def pick_contact(nama_keluarga, wa_raw, email_raw):
    """Pilih kontak kepala keluarga (nama pertama di Nama Keluarga); jika tidak ada, ASN pertama."""
    kk = str(nama_keluarga).split("/")[0].strip().upper()
    emails, phones = parse_multi(email_raw), parse_multi(wa_raw)
    names = [n for n, _ in emails + phones if n]
    chosen = kk if kk in names else (names[0] if names else None)

    def get(pairs):
        for n, v in pairs:
            if n == chosen or n is None:
                return v
        return pairs[0][1] if pairs else ""

    return (chosen or ""), get(emails), get(phones)


def build_targets():
    df = pd.read_excel(SRC_XLSX, sheet_name=SHEET, dtype=str)
    rows, skipped = [], []
    for _, r in df.iterrows():
        base = {"no": r["No"], "assignment_id": r["Assignment ID"], "nama_keluarga": str(r["Nama Keluarga"]).strip(),
                "alamat": r["Alamat"] if isinstance(r.get("Alamat"), str) else ""}
        if pd.isna(r["Email"]) and pd.isna(r["No. WhatsApp"]):
            continue
        if str(r.get("Keberadaan Keluarga di Fasih-SM", "")).strip().upper() == "DITEMUKAN/BARU":
            skipped.append({**base, "alasan": "Keberadaan Keluarga: Ditemukan/Baru"})
            continue
        ket = r.get("Keterangan")
        if isinstance(ket, str) and EXCLUDE_KETERANGAN.search(ket):
            skipped.append({**base, "alasan": f"Keterangan: {ket}"})
            continue
        asn, email, phone = pick_contact(r["Nama Keluarga"], r["No. WhatsApp"], r["Email"])
        email, phone = email.strip(), re.sub(r"\D", "", phone)
        note = []
        if phone and not PHONE_RE.match(phone):
            note.append(f"nomor tidak valid ({phone}) -> dikosongkan")
            phone = ""
        if not EMAIL_RE.match(email):
            skipped.append({**base, "alasan": f"email kosong/tidak valid ({email})"})
            continue
        rows.append({**base, "asn_dipilih": asn, "email": email, "telepon": phone, "catatan": "; ".join(note)})
    return pd.DataFrame(rows), pd.DataFrame(skipped)


# ----------------------------------------------------------------------------- helpers
def pause(rng):
    time.sleep(random.uniform(*rng))


def hclick(locator, **kw):
    """Klik dengan hover dulu, jeda singkat, dan lama tekan mouse yang acak."""
    try:
        locator.hover(timeout=3000, **({"force": True} if kw.get("force") else {}))
    except Exception:
        pass
    pause(HOVER_DELAY)
    locator.click(delay=random.randint(*CLICK_HOLD_MS), **kw)


def paste_text(locator, text):
    """Klik field, pilih semua isi lama, lalu tempel teks sekaligus (seperti Ctrl+V)."""
    hclick(locator)
    pause((0.1, 0.4))
    locator.press("Control+a")
    pause((0.1, 0.4))
    locator.page.keyboard.insert_text(text)
    if locator.input_value() != text:  # cadangan jika tempel tidak terbaca oleh form
        locator.fill(text)


def norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip().upper()


def first_visible(*locators, timeout=8000):
    """Kembalikan locator pertama yang terlihat dari beberapa kandidat selector."""
    deadline = time.time() + timeout / 1000
    while time.time() < deadline:
        for loc in locators:
            try:
                if loc.count() and loc.first.is_visible():
                    return loc.first
            except Exception:
                pass
        time.sleep(0.25)
    return None


class Step(Exception):
    """Kegagalan pada satu langkah; status = kode untuk log."""
    def __init__(self, status, msg):
        super().__init__(msg)
        self.status = status


class Bot:
    def __init__(self, page, execute=False, recon=False):
        self.page, self.execute, self.recon = page, execute, recon
        self.recon_i = 0

    def dump(self, label):
        if not self.recon:
            return
        RECON_DIR.mkdir(exist_ok=True)
        self.recon_i += 1
        stem = RECON_DIR / f"{self.recon_i:02d}_{label}"
        self.page.screenshot(path=f"{stem}.png", full_page=True)
        stem.with_suffix(".html").write_text(self.page.content(), encoding="utf-8")
        print(f"   [recon] {stem.name}")

    def wait_idle(self):
        try:
            self.page.wait_for_load_state("networkidle", timeout=10000)
        except PWTimeout:
            pass
        spinner = self.page.locator(".spinner, .loading, .p-progress-spinner, .ngx-spinner-overlay, [role=progressbar]")
        try:
            spinner.first.wait_for(state="hidden", timeout=10000)
        except Exception:
            pass

    # --- langkah 1: cari & temukan baris OPEN
    def search(self, name):
        # placeholder persis "Cari..." (bukan "Cari periode..." di sidebar kiri)
        box = first_visible(self.page.get_by_placeholder("Cari...", exact=True))
        if not box:
            raise Step("ERROR_SELECTOR", "kolom search tidak ditemukan")
        paste_text(box, name)
        pause((0.3, 0.7))
        box.press("Enter")
        pause(STEP_DELAY)
        self.wait_idle()
        self.dump("hasil_search")

    def find_open_row(self, name, alamat):
        target = norm(name)
        rows = self.page.locator("table tbody tr")
        deadline = time.time() + 15
        while time.time() < deadline:
            texts = rows.all_inner_texts()
            if any(target in norm(t) for t in texts):
                break
            time.sleep(0.5)
        texts = rows.all_inner_texts()
        headers = [norm(h) for h in self.page.locator("table thead th").all_inner_texts()]
        i_nama = next((i for i, h in enumerate(headers) if h.startswith("NAMA KELUARGA")), None)
        i_status = headers.index("STATUS") if "STATUS" in headers else None
        i_mode = headers.index("MODE") if "MODE" in headers else None
        if i_nama is None or i_status is None:
            raise Step("ERROR_SELECTOR", f"kolom Nama Keluarga/Status tidak ditemukan di header: {headers}")
        cells = [[norm(c) for c in rows.nth(i).locator("td").all_inner_texts()] for i in range(len(texts))]
        cand = [i for i, c in enumerate(cells) if len(c) > i_status and c[i_nama] == target]
        if not cand:
            tampil = " | ".join(c[i_nama] for c in cells if len(c) > i_nama)[:150]
            raise Step("SKIP_TIDAK_DITEMUKAN", f"tidak ada baris bernama persis '{name}' "
                                               f"({len(texts)} baris tampil: {tampil or '-'})")
        if i_mode is not None:
            sudah = [i for i in cand if "CAWI" in cells[i][i_mode]]
            if sudah and len(sudah) == len(cand):
                raise Step("SKIP_SUDAH_CAWI", "mode assignment sudah CAWI (kemungkinan sudah pernah dikirim)")
            cand = [i for i in cand if i not in sudah]
        open_rows = [i for i in cand if re.search(r"\bOPEN\b", cells[i][i_status])]
        if not open_rows:
            raise Step("SKIP_BUKAN_OPEN", "status bukan OPEN: " + " | ".join(cells[i][i_status] for i in cand))
        if len(open_rows) > 1 and alamat:
            by_alamat = [i for i in open_rows if norm(alamat) in norm(texts[i])]
            open_rows = by_alamat or open_rows
        if len(open_rows) > 1:
            raise Step("SKIP_AMBIGU", f"{len(open_rows)} baris OPEN bernama sama")
        return rows.nth(open_rows[0])

    # --- langkah 2: titik 3 -> Ganti Mode
    def open_ganti_mode(self, row):
        row.scroll_into_view_if_needed()
        last = row.locator("td").last
        kebab = first_visible(last.locator("button"), last.locator("[role=button]"), last.locator("a"),
                              last.locator("i, svg, span"), timeout=4000)
        if not kebab:
            raise Step("ERROR_SELECTOR", "tombol titik 3 tidak ditemukan")
        hclick(kebab)
        pause(STEP_DELAY)
        self.dump("menu_titik3")
        item = first_visible(self.page.get_by_role("menuitem", name=re.compile("Ganti Mode", re.I)),
                             self.page.get_by_text(re.compile(r"^\s*Ganti Mode\s*$", re.I)))
        if not item:
            raise Step("ERROR_SELECTOR", "menu 'Ganti Mode' tidak muncul")
        hclick(item)
        pause(STEP_DELAY)
        dlg = first_visible(
            self.page.locator("[role=dialog]").filter(has_text=re.compile("Mode Pendataan", re.I)),
            self.page.locator(".modal, .p-dialog, .ant-modal, mat-dialog-container, .v-dialog")
                .filter(has_text=re.compile("Mode Pendataan", re.I)),
        )
        if not dlg:
            raise Step("ERROR_SELECTOR", "dialog Ganti Mode tidak muncul")
        self.dump("dialog_ganti_mode")
        return dlg

    # --- langkah 3: isi form
    def set_cawi(self, dlg):
        sel = dlg.locator("select")
        if sel.count():
            for i in range(sel.count()):
                opts = sel.nth(i).locator("option").all_inner_texts()
                lbl = next((o for o in opts if "CAWI" in o.upper()), None)
                if lbl:
                    sel.nth(i).select_option(label=lbl)
                    return
        radio = dlg.get_by_label(re.compile(r"^\s*CAWI", re.I))
        if radio.count() and radio.first.is_visible():
            radio.first.check()
            return
        trigger = first_visible(
            dlg.get_by_role("combobox"),
            dlg.get_by_text(re.compile("Mode Pendataan", re.I)).locator(
                "xpath=following::*[@role='combobox' or contains(@class,'select') or contains(@class,'dropdown')][1]"),
            timeout=4000)
        if not trigger:
            raise Step("ERROR_SELECTOR", "dropdown Mode Pendataan tidak ditemukan")
        hclick(trigger)
        pause((0.4, 1.0))
        opt = first_visible(self.page.get_by_role("option", name=re.compile(r"CAWI", re.I)),
                            self.page.get_by_text(re.compile(r"^\s*CAWI\s*$", re.I)), timeout=4000)
        if not opt:
            raise Step("ERROR_SELECTOR", "opsi CAWI tidak ditemukan")
        hclick(opt)

    def fill_form(self, dlg, email, phone):
        self.set_cawi(dlg)
        pause(STEP_DELAY)
        self.dump("cawi_dipilih")
        email_in = first_visible(
            dlg.get_by_label(re.compile("email", re.I)),
            dlg.locator("input[type=email], input[placeholder*=mail i], input[name*=mail i], "
                        "input[formcontrolname*=mail i], input[id*=mail i]"))
        if not email_in:
            raise Step("ERROR_SELECTOR", "field Email Responden tidak ditemukan")
        paste_text(email_in, email)
        pause(STEP_DELAY)
        if phone:
            phone_in = first_visible(
                dlg.get_by_label(re.compile(r"telepon|telp|phone|\bhp\b|whatsapp", re.I)),
                dlg.locator("input[type=tel], input[placeholder*=telepon i], input[placeholder*=phone i], "
                            "input[name*=phone i], input[formcontrolname*=phone i], input[id*=phone i], "
                            "input[name*=telp i], input[formcontrolname*=telp i]"))
            if not phone_in:
                raise Step("ERROR_SELECTOR", "field Nomor Telepon Responden tidak ditemukan")
            paste_text(phone_in, phone)
            pause(STEP_DELAY)
        else:
            # baris email-saja: nomor & switch Kirim Pesan (WhatsApp) dibiarkan apa adanya
            self.dump("form_terisi")
            return
        sw = first_visible(
            dlg.get_by_role("switch", name=re.compile("Kirim Pesan", re.I)),
            dlg.get_by_label(re.compile("Kirim Pesan", re.I)),
            dlg.get_by_text(re.compile("Kirim Pesan", re.I)).locator(
                "xpath=(ancestor-or-self::*[.//input[@type='checkbox'] or .//*[@role='switch']][1]"
                "//*[self::input[@type='checkbox'] or @role='switch'])[1]"),
            dlg.get_by_role("switch"),
            timeout=4000)
        if not sw:
            raise Step("ERROR_SELECTOR", "switch Kirim Pesan tidak ditemukan")
        if not self.is_on(sw):
            hclick(sw)
            pause((0.3, 0.8))
        if not self.is_on(sw):
            raise Step("ERROR_SELECTOR", "switch Kirim Pesan gagal dinyalakan")
        self.dump("form_terisi")

    @staticmethod
    def is_on(loc):
        try:
            return loc.is_checked()
        except Exception:
            return (loc.get_attribute("aria-checked") == "true"
                    or "checked" in (loc.get_attribute("class") or "").lower())

    # --- langkah 4: submit / batal
    def close_dialog(self, dlg):
        btn = first_visible(dlg.get_by_role("button", name=re.compile(r"batal|cancel|tutup|close", re.I)), timeout=2000)
        if btn:
            hclick(btn)
        else:
            self.page.keyboard.press("Escape")
        try:
            dlg.wait_for(state="hidden", timeout=5000)
        except PWTimeout:
            self.page.keyboard.press("Escape")

    def submit(self, dlg):
        btn = first_visible(dlg.get_by_role("button", name=re.compile("Ubah Mode Pendataan", re.I)), timeout=4000)
        if not btn:
            raise Step("ERROR_SELECTOR", "tombol 'Ubah Mode Pendataan' tidak ditemukan")
        if not btn.is_enabled():
            raise Step("ERROR_FORM", "tombol 'Ubah Mode Pendataan' nonaktif (validasi form gagal?)")
        if not self.execute:
            return "DRYRUN_OK", "berhenti sebelum klik Ubah Mode Pendataan"
        hclick(btn)
        pause((0.7, 1.5))
        confirm = first_visible(
            self.page.locator("[role=dialog], .swal2-popup, .modal").filter(
                has_text=re.compile(r"yakin|konfirmasi", re.I)).get_by_role(
                "button", name=re.compile(r"^\s*(ya|yakin|ok|lanjut)", re.I)), timeout=2500)
        if confirm:
            hclick(confirm)
        msg = ""
        toast = first_visible(self.page.locator("[role=alert], .toast, .p-toast-message, .swal2-popup, "
                                                ".ant-message, .snackbar, .mat-snack-bar-container"), timeout=6000)
        if toast:
            msg = norm(toast.inner_text())[:200]
        try:
            dlg.wait_for(state="hidden", timeout=10000)
        except PWTimeout:
            raise Step("ERROR_SUBMIT", f"dialog tidak tertutup setelah submit. {msg}")
        if re.search(r"gagal|error|failed", msg, re.I):
            raise Step("ERROR_SUBMIT", msg)
        return "SUCCESS", msg

    def session_ok(self):
        if not re.search(r"fasih-sm\.bps\.go\.id/app/surveys/.+/data", self.page.url):
            return False
        return not self.page.get_by_text(re.compile(r"Login SSO", re.I)).count()

    def ensure_session(self):
        """Jika sesi habis (halaman login), tunggu Anda login ulang lalu lanjut."""
        if self.session_ok():
            return
        print("\n!!!! SESI FASIH HABIS - silakan login ulang di jendela Chrome ini. Bot menunggu (maks 30 menit)...")
        try:
            import winsound
            for _ in range(3):
                winsound.Beep(1000, 400)
        except Exception:
            pass
        wait_login(self.page, timeout_min=30)
        pause(STEP_DELAY)

    def reload_data(self):
        self.page.goto(DATA_URL)
        self.ensure_session()
        self.page.locator("table").first.wait_for(timeout=60_000)
        self.wait_idle()

    def process(self, t):
        self.ensure_session()
        try:
            self.search(t["nama_keluarga"])
            row = self.find_open_row(t["nama_keluarga"], t["alamat"])
        except Step as e:
            if e.status not in ("SKIP_TIDAK_DITEMUKAN", "ERROR_SELECTOR"):
                raise
            # bisa jadi hasil search belum ter-refresh / sesi habis: muat ulang halaman lalu coba sekali lagi
            print(f"    .. {e.status}, muat ulang halaman & coba lagi")
            self.reload_data()
            self.search(t["nama_keluarga"])
            row = self.find_open_row(t["nama_keluarga"], t["alamat"])
        dlg = self.open_ganti_mode(row)
        try:
            self.fill_form(dlg, t["email"], t["telepon"])
            status, msg = self.submit(dlg)
            if status == "DRYRUN_OK":
                SHOT_DIR.mkdir(exist_ok=True)
                dlg.screenshot(path=str(SHOT_DIR / f"{t['no']}_{re.sub(r'\W+', '_', t['nama_keluarga'])[:40]}.png"))
                self.close_dialog(dlg)
            return status, msg
        except Step:
            self.close_dialog(dlg)
            raise


# ----------------------------------------------------------------------------- main
def wait_login(page, timeout_min=5):
    if "/data" not in page.url or "fasih-sm.bps.go.id" not in page.url:
        page.goto(DATA_URL)
    print(f">> Silakan login SSO BPS di jendela Chrome jika diminta (maks {timeout_min} menit)...")
    page.wait_for_url(re.compile(r"fasih-sm\.bps\.go\.id/app/surveys/.+/data"), timeout=timeout_min * 60_000)
    page.locator("table").first.wait_for(timeout=120_000)
    print(">> Halaman data siap.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--execute", action="store_true", help="benar-benar klik 'Ubah Mode Pendataan'")
    ap.add_argument("--recon", action="store_true", help="dry-run 1 assignment, simpan HTML+screenshot tiap langkah")
    ap.add_argument("--prepare-only", action="store_true", help="hanya buat daftar_target.csv")
    ap.add_argument("--limit", type=int, default=0, help="maks assignment diproses (0 = semua)")
    ap.add_argument("--only", nargs="*", help="proses hanya Nama Keluarga ini")
    ap.add_argument("--port", type=int, default=9222, help="port remote debugging Chrome (default 9222)")
    ap.add_argument("--launch", action="store_true", help="mode lama: Playwright membuka Chrome sendiri")
    ap.add_argument("--part", help="bagian K dari N, mis. 1/2 dan 2/2 untuk 2 window paralel")
    ap.add_argument("--exclude-part", help="lewati target milik bagian K/N (mis. 1/4 yang masih berjalan di window lain)")
    args = ap.parse_args()
    if args.recon:
        args.execute, args.limit = False, 1
    part_k = part_n = None
    if args.part:
        m = re.fullmatch(r"(\d+)/(\d+)", args.part)
        if not m or not 1 <= int(m[1]) <= int(m[2]):
            ap.error("--part harus berformat K/N, mis. 1/2")
        part_k, part_n = int(m[1]), int(m[2])
    excl = None
    if args.exclude_part:
        m = re.fullmatch(r"(\d+)/(\d+)", args.exclude_part)
        if not m or not 1 <= int(m[1]) <= int(m[2]):
            ap.error("--exclude-part harus berformat K/N, mis. 1/4")
        excl = (int(m[1]), int(m[2]))

    targets, skipped = build_targets()
    if not args.part:  # hindari 2 proses menulis file yang sama bersamaan
        targets.to_csv(BASE / "daftar_target.csv", index=False, encoding="utf-8-sig")
        skipped.to_csv(BASE / "daftar_dilewati.csv", index=False, encoding="utf-8-sig")
    print(f"Target: {len(targets)} | dilewati saat persiapan: {len(skipped)}")
    if args.prepare_only:
        return

    LOG_DIR.mkdir(exist_ok=True)
    mode = "execute" if args.execute else "dryrun"
    log_path = LOG_DIR / (f"log_{mode}_part{part_k}.csv" if part_k else f"log_{mode}.csv")
    # status selesai dibaca dari SEMUA log mode ini (termasuk bagian lain), supaya tidak dobel kirim
    done = set()
    ok = "SUCCESS" if args.execute else "DRYRUN_OK"
    for f in LOG_DIR.glob(f"log_{mode}*.csv"):
        prev = pd.read_csv(f, dtype=str)
        done |= set(prev.loc[prev["status"].isin([ok, "SKIP_BUKAN_OPEN", "SKIP_SUDAH_CAWI"]), "assignment_id"])
    if excl:  # dihitung dari urutan daftar_target penuh, sama seperti --part
        n0 = len(targets)
        targets = targets[[i % excl[1] != excl[0] - 1 for i in range(len(targets))]]
        print(f"Kecualikan bagian {excl[0]}/{excl[1]}: {n0 - len(targets)} target dilewati")
    if part_k:
        targets = targets.iloc[part_k - 1::part_n]  # selang-seling: baris ke-1,3,5.. / 2,4,6..
        print(f"Bagian {part_k}/{part_n}: {len(targets)} target")
    todo = targets[~targets["assignment_id"].isin(done)]
    if args.only:
        todo = todo[todo["nama_keluarga"].str.upper().isin([n.upper() for n in args.only])]
    if args.limit:
        todo = todo.head(args.limit)
    print(f"Akan diproses: {len(todo)} (sudah selesai sebelumnya: {len(done)}) | "
          f"MODE: {'EXECUTE' if args.execute else 'DRY-RUN'} | port {args.port}")
    if args.execute:
        print("!! MODE EXECUTE: pesan CAWI akan dikirim ke responden. Ctrl+C untuk berhenti.")

    new_log = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8-sig") as fh, sync_playwright() as p:
        w = csv.writer(fh)
        if new_log:
            w.writerow(["waktu", "no", "assignment_id", "nama_keluarga", "email", "telepon", "status", "pesan"])
        ctx = None
        if args.launch:
            ctx = p.chromium.launch_persistent_context(
                str(PROFILE_DIR), channel="chrome", headless=False, viewport={"width": 1440, "height": 900},
                slow_mo=random.randint(40, 90))
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
        else:
            try:
                browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}",
                                                      slow_mo=random.randint(40, 90))
            except Exception as e:
                print(f"!! Tidak bisa terhubung ke Chrome di port {args.port}. Jalankan buka_chrome.bat dulu.\n   {e}")
                return 1
            pages = [pg for c in browser.contexts for pg in c.pages]
            page = next((pg for pg in pages if "fasih-sm.bps.go.id" in pg.url), None) \
                or (pages[0] if pages else browser.contexts[0].new_page())
            page.bring_to_front()
        page.set_default_timeout(15000)
        wait_login(page)
        bot = Bot(page, execute=args.execute, recon=args.recon)
        next_break = random.randint(*LONG_BREAK_EVERY)
        try:
            for n, t in enumerate(todo.to_dict("records"), 1):
                if (BASE / "STOP").exists():
                    print(">> File STOP ditemukan, berhenti.")
                    break
                print(f"[{n}/{len(todo)}] {t['nama_keluarga']} <{t['email']}> {t['telepon']}")
                try:
                    status, msg = bot.process(t)
                except Step as e:
                    status, msg = e.status, str(e)
                except Exception as e:
                    status, msg = "ERROR", f"{type(e).__name__}: {e}"[:300]
                if status.startswith("ERROR"):
                    SHOT_DIR.mkdir(exist_ok=True)
                    try:
                        page.screenshot(path=str(SHOT_DIR / f"ERROR_{t['no']}.png"))
                        page.keyboard.press("Escape")
                    except Exception:
                        pass
                print(f"    -> {status} {msg}")
                w.writerow([datetime.now().isoformat(timespec="seconds"), t["no"], t["assignment_id"],
                            t["nama_keluarga"], t["email"], t["telepon"], status, msg])
                fh.flush()
                if n == len(todo):
                    break
                if n >= next_break:
                    d = random.uniform(*LONG_BREAK)
                    print(f">> istirahat {d:.0f} detik")
                    time.sleep(d)
                    next_break = n + random.randint(*LONG_BREAK_EVERY)
                else:
                    pause(ITEM_DELAY)
        except KeyboardInterrupt:
            print("\n>> Dihentikan oleh pengguna.")
        finally:
            if args.recon:
                print(f">> File recon ada di {RECON_DIR}")
            if ctx:
                ctx.close()  # mode --launch; mode CDP: Chrome Anda tetap terbuka


if __name__ == "__main__":
    sys.exit(main())
