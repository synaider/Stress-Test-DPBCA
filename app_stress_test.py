"""
Aplikasi GUI — Stress Test Risiko (Permintaan OJK: Risiko Pasar, Kredit, Likuiditas)
Dana Pensiun BCA · Manajemen Risiko

Menjalankan:
    pip install -r requirements.txt
    streamlit run app_stress_test.py

Hasil akhir: file Excel "Stress Test Risiko <Bulan> <Tahun>.xlsx" dengan format
sama persis seperti file "Permintaan OJK Datastress test risiko pasar kredit
likuiditas", tersimpan di folder hasil/ dan bisa diunduh dari aplikasi.
"""
from __future__ import annotations

import datetime as dt
import io
import math
from dataclasses import asdict

import pandas as pd
import streamlit as st

import ffs_reader as fr
import stress_core as sc

st.set_page_config(page_title="Stress Test Risiko · DPBCA", page_icon="📊", layout="wide")
S = st.session_state


def _w(fn, *args, **kwargs):
    """Lebar penuh, kompatibel Streamlit lama (use_container_width) & baru (width='stretch')."""
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception as e:  # versi lama: parameter width belum ada / hanya angka
        if isinstance(e, TypeError) or "width" in str(e).lower():
            return fn(*args, use_container_width=True, **kwargs)
        raise

# --------------------------------------------------------------------------
# Gaya tampilan
# --------------------------------------------------------------------------
st.markdown("""
<style>
:root { --dp-blue:#1B4A8F; --dp-blue-2:#2F6FD0; --dp-soft:#EEF3FB; }
.block-container { padding-top: 3.2rem; max-width: 1350px; }
h1, h2, h3 { color: var(--dp-blue); }
.dp-head { display:flex; align-items:center; gap:14px; margin-bottom:.4rem; }
.dp-head .t1 { font-size:1.55rem; font-weight:700; color:var(--dp-blue); line-height:1.1; }
.dp-head .t2 { font-size:.92rem; color:#5b6575; }
.dp-card { background:var(--dp-soft); border-radius:10px; padding:.7rem .9rem; border-left:4px solid var(--dp-blue); }
.dp-ok { color:#137333; font-weight:600; } .dp-bad { color:#B3261E; font-weight:600; }
div[data-testid="stMetricValue"] { font-size:1.35rem; }
.small-note { font-size:.82rem; color:#5b6575; }
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# Daftar field input (widget key = "f_" + nama field)
# --------------------------------------------------------------------------
NUM_FIELDS = ["ihsg_existing", "ihsg_escalated", "ihsg_severe",
              "kurs_existing", "kurs_baseline", "kurs_escalated", "kurs_severe",
              "ex_total_aset", "ex_spi", "ex_laba", "ex_delta_spi",
              "pr_total_aset", "pr_spi", "pr_laba", "pr_delta_spi", "aset_neto",
              "yc_nilai_wajar", "yc_delta_spi", "spi_saham_des_juta", "spi_saham_rep_juta",
              "rd_saham_np", "rd_saham_nw", "rd_campuran_np", "rd_campuran_nw",
              "likuiditas_eksposur"]
INT_FIELDS = ["yc_tenor", "likuiditas_tenor"]
PCT_FIELDS = ["yield_existing", "yield_escalated", "yield_severe"]
KELAS_OPS = ["", "AAA", "AA", "A", "D"]
MI_OPS = ["", "Sangat Bagus", "Bagus", "Cukup"]
OTHER_FUND = "➕ Reksa dana lain (ketik nama)"


def fmt_id(x, nd=2):
    if x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))):
        return "–"
    s = f"{x:,.{nd}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def pct_id(x, nd=2):
    return "–" if x is None else fmt_id(x * 100, nd) + "%"


# --------------------------------------------------------------------------
# Sinkronisasi StressInputs <-> session_state
# --------------------------------------------------------------------------
def put_inputs(inp: sc.StressInputs):
    for f in NUM_FIELDS:
        S["f_" + f] = float(getattr(inp, f) or 0.0)
    for f in INT_FIELDS:
        S["f_" + f] = int(getattr(inp, f) or 10)
    for f in PCT_FIELDS:
        S["f_" + f] = round(float(getattr(inp, f) or 0.0) * 100, 8)
    S.f_tahun = int(inp.tahun)
    S.f_bulan = sc.BULAN[int(inp.bulan)]
    S.f_lgd = round(float(inp.lgd) * 100, 6)
    for key in ("esc", "sev"):
        S[f"f_b_manual_{key}"] = bool(getattr(inp, f"b_manual_{key}", False))
        S[f"f_b_{key}"] = round(float(getattr(inp, f"b_{key}", 0.0)) * 100, 8)
    S.pd_kelas = dict(inp.pd_kelas or sc.DEFAULT_PD_KELAS)
    S.pd_mi = dict(inp.pd_mi or sc.DEFAULT_PD_MI)
    set_funds([asdict(f) for f in inp.funds])
    set_credit([asdict(c) for c in inp.credit])
    set_mi([asdict(m) for m in inp.mi])


def set_funds(funds: list):
    # hapus key slot lama
    for k in [k for k in list(S.keys()) if str(k).startswith("fund_")]:
        del S[k]
    S.n_funds = max(1, len(funds))
    for i, f in enumerate(funds):
        S[f"fund_{i}_pilih"] = f.get("nama", "")
        S[f"fund_{i}_nama"] = f.get("nama", "")
        S[f"fund_{i}_kode"] = f.get("kode_raw", "") or f.get("nama", "")
        S[f"fund_{i}_saham"] = float(f.get("saham", 0.0))
        S[f"fund_{i}_obl"] = float(f.get("obligasi", 0.0))
        S[f"fund_{i}_pu"] = float(f.get("pasar_uang", 0.0))
        S[f"fund_{i}_nw"] = float(f.get("nilai_wajar", 0.0))
        S[f"fund_{i}_sumber"] = f.get("sumber", "")


def set_credit(rows: list):
    df = pd.DataFrame(rows, columns=["nama", "eksposur", "rating", "kelas", "pd"])
    df = df.rename(columns={"nama": "Grup / Pihak", "eksposur": "Eksposur (miliar)",
                            "rating": "Rating", "kelas": "Kelas"})[
        ["Grup / Pihak", "Eksposur (miliar)", "Rating", "Kelas"]]
    df["Kelas"] = df["Kelas"].fillna("").astype(str)
    S.credit_df = df.reset_index(drop=True)
    S.credit_edited = S.credit_df
    S.credit_ver = S.get("credit_ver", 0) + 1


def set_mi(rows: list):
    df = pd.DataFrame(rows, columns=["pihak", "nama", "nominal", "rating"])
    df = df.rename(columns={"pihak": "Pihak (RAW)", "nama": "Nama Manajer Investasi",
                            "nominal": "Nominal (miliar)", "rating": "Rating"})
    S.mi_df = df.reset_index(drop=True)
    S.mi_edited = S.mi_df
    S.mi_ver = S.get("mi_ver", 0) + 1


def _v(x, default=None):
    """Nilai sel data_editor: NaN/None -> default."""
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)) or pd.isna(x):
            return default
    except (TypeError, ValueError):
        pass
    return x


def collect_inputs() -> sc.StressInputs:
    inp = sc.StressInputs()
    inp.tahun = int(S.f_tahun)
    inp.bulan = sc.BULAN.index(S.f_bulan)
    for f in NUM_FIELDS:
        setattr(inp, f, float(S.get("f_" + f) or 0.0))
    for f in INT_FIELDS:
        setattr(inp, f, int(S.get("f_" + f) or 10))
    for f in PCT_FIELDS:
        setattr(inp, f, round(float(S.get("f_" + f) or 0.0) / 100, 10))
    inp.lgd = round(float(S.get("f_lgd", 45.0)) / 100, 8)
    inp.pd_kelas = dict(S.pd_kelas)
    inp.pd_mi = dict(S.pd_mi)
    for key in ("esc", "sev"):
        setattr(inp, f"b_manual_{key}", bool(S.get(f"f_b_manual_{key}", False)))
        setattr(inp, f"b_{key}", round(float(S.get(f"f_b_{key}") or 0.0) / 100, 10))
    funds = []
    for i in range(S.n_funds):
        funds.append(sc.FundSlot(nama=str(S.get(f"fund_{i}_nama", "")).strip(),
                                 kode_raw=str(S.get(f"fund_{i}_kode", "")),
                                 saham=float(S.get(f"fund_{i}_saham", 0.0)),
                                 obligasi=float(S.get(f"fund_{i}_obl", 0.0)),
                                 pasar_uang=float(S.get(f"fund_{i}_pu", 0.0)),
                                 nilai_wajar=float(S.get(f"fund_{i}_nw", 0.0)),
                                 sumber=str(S.get(f"fund_{i}_sumber", ""))))
    inp.funds = funds
    cdf = S.get("credit_edited", S.credit_df)
    credit = []
    for _, r in cdf.iterrows():
        nama = str(_v(r.get("Grup / Pihak"), "")).strip()
        if not nama:
            continue
        rating = str(_v(r.get("Rating"), "")).strip()
        kelas = str(_v(r.get("Kelas"), "")).strip() or (sc.rating_class(rating) or "")
        eks = float(_v(r.get("Eksposur (miliar)"), 0.0))
        credit.append(sc.CreditRow(nama=nama, eksposur=eks, rating=rating, kelas=kelas,
                                   pd=float(inp.pd_kelas.get(kelas, 0.0)) if kelas else 0.0))
    inp.credit = credit
    mdf = S.get("mi_edited", S.mi_df)
    mi = []
    for _, r in mdf.iterrows():
        nama = str(_v(r.get("Nama Manajer Investasi"), "")).strip()
        if not nama:
            continue
        mi.append(sc.MIRow(pihak=str(_v(r.get("Pihak (RAW)"), nama) or nama).strip(), nama=nama,
                           nominal=float(_v(r.get("Nominal (miliar)"), 0.0)),
                           rating=str(_v(r.get("Rating"), "")).strip()))
    inp.mi = mi
    return inp


# --------------------------------------------------------------------------
# Inisialisasi
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _template_defaults_dict(path_str: str, mtime: float) -> dict:
    return asdict(sc.defaults_from_template(path_str))


def init_state(force_template=False):
    tpl = S.get("template_path", str(sc.TEMPLATE_PATH))
    inp = None if force_template else sc.load_last_inputs()
    src = "input terakhir yang disimpan"
    if inp is None:
        inp = sc.StressInputs.from_dict(_template_defaults_dict(tpl, sc.Path(tpl).stat().st_mtime))
        src = "template (file bulan sebelumnya)"
    put_inputs(inp)
    S.defaults_source = src
    S.master = sc.load_master(tpl)
    S.catalog = fr.load_catalog()
    S.raw_headers = sc.template_raw_headers(tpl)
    S.init_done = True


if "template_path" not in S:
    S.template_path = str(sc.TEMPLATE_PATH)
if not S.get("init_done"):
    S.raw = None
    S.budget = None
    S.result = None
    S.msgs = []
    init_state()


def flash(kind: str, text: str):
    S.msgs = S.get("msgs", []) + [(kind, text)]


# --------------------------------------------------------------------------
# Callback (dijalankan sebelum halaman digambar ulang)
# --------------------------------------------------------------------------
def on_raw_upload():
    up = S.get("up_raw")
    if up is None:
        return
    try:
        raw = sc.read_rekap(io.BytesIO(up.getvalue()), S.raw_headers, up.name)
    except Exception as e:
        S.raw = None
        flash("error", f"Gagal membaca Rekap Investasi: {e}")
        return
    S.raw = raw
    S.result = None
    if raw.period:
        S.f_tahun, S.f_bulan = raw.period.year, sc.BULAN[raw.period.month]
    refresh_from_raw(raw, quiet=True)
    flash("success", f"Rekap Investasi terbaca: {len(raw.rows)} baris dari sheet '{raw.sheet_name}'"
          + (f", periode {sc.BULAN[raw.period.month]} {raw.period.year}." if raw.period else "."))


def refresh_from_raw(raw=None, quiet=False):
    raw = raw or S.raw
    if raw is None:
        return
    pdk = dict(S.pd_kelas)
    set_credit([asdict(c) for c in sc.build_credit_rows(raw, S.master, pdk)])
    set_mi([asdict(m) for m in sc.build_mi_rows(raw, S.master)])
    for i in range(S.n_funds):
        key = S.get(f"fund_{i}_kode") or S.get(f"fund_{i}_nama")
        nm, nw = sc.find_fund_nw(raw, key)
        if nw:
            S[f"fund_{i}_nw"] = float(nw)
            S[f"fund_{i}_kode"] = nm
    S.f_spi_saham_rep_juta = round(sc.spi_saham_existing_juta(raw), 6)
    inp = collect_inputs()
    S.f_yc_nilai_wajar = sc.round_half_up(sc.total_terdampak_suku_bunga(raw, inp))
    if not quiet:
        flash("success", "Data kredit, MI, Nilai Wajar reksa dana, dan Nilai Wajar kalkulator yield "
                         "diperbarui dari RAW.")


def on_budget_fill():
    up = S.get("up_budget")
    if up is None:
        flash("warning", "Pilih file Budget & Realisasi terlebih dahulu.")
        return
    try:
        bd = sc.read_budget(io.BytesIO(up.getvalue()), int(S.f_tahun), sc.BULAN.index(S.f_bulan))
    except Exception as e:
        flash("error", f"Gagal membaca file Budget: {e}")
        return
    S.budget = bd
    n = 0
    for k, v in bd.values.items():
        if ("f_" + k) in S:
            S["f_" + k] = float(v)
            n += 1
    flash("success" if not bd.warnings else "warning",
          f"{n} angka terisi dari Budget & Realisasi." + (" Ada catatan, lihat di bawah." if bd.warnings else ""))


def on_rd_from_raw():
    if S.raw is None:
        flash("warning", "Upload Rekap Investasi dulu.")
        return
    per = sc.spi_rd_per_jenis(S.raw)
    sah = per.get("Saham", (0.0, 0.0))
    cam = per.get("Campuran", (0.0, 0.0))
    S.f_rd_saham_np, S.f_rd_saham_nw = float(sah[0]), float(sah[1])
    S.f_rd_campuran_np, S.f_rd_campuran_nw = float(cam[0]), float(cam[1])
    flash("info", "SPI reksa dana diisi dari RAW. Cocokkan dengan Laporan SPI Reksadana per Jenis (DISK).")


def on_fund_select(i: int):
    pilih = S.get(f"fund_{i}_pilih", "")
    if pilih == OTHER_FUND:
        S[f"fund_{i}_nama"] = ""
        S[f"fund_{i}_kode"] = ""
        S[f"fund_{i}_nw"] = 0.0
        S[f"fund_{i}_saham"] = S[f"fund_{i}_obl"] = S[f"fund_{i}_pu"] = 0.0
        S[f"fund_{i}_sumber"] = ""
        return
    S[f"fund_{i}_nama"] = sc.smart_title(pilih) if pilih.isupper() else pilih
    S[f"fund_{i}_kode"] = pilih
    S[f"fund_{i}_sumber"] = ""
    S[f"fund_{i}_log"] = []
    # komposisi reksa dana sebelumnya tidak berlaku lagi -> kosongkan, wajib diisi ulang
    S[f"fund_{i}_saham"] = S[f"fund_{i}_obl"] = S[f"fund_{i}_pu"] = 0.0
    flash("info", f"Reksa dana #{i + 1} diganti ke '{S[f'fund_{i}_nama']}'. Komposisi dikosongkan — tekan "
                  "'Ambil FFS terbaru', upload PDF FFS, atau isi manual.")
    if S.raw is not None:
        nm, nw = sc.find_fund_nw(S.raw, pilih)
        S[f"fund_{i}_nw"] = float(nw)
        if nm:
            S[f"fund_{i}_kode"] = nm


def _fund_entry(i: int) -> dict:
    nama = S.get(f"fund_{i}_nama", "")
    e = fr.find_in_catalog(S.catalog, nama)
    if e is None:
        mi = ""
        if S.raw is not None:
            for rd in sc.reksa_dana_list(S.raw):
                if sc._norm(rd["nama"]) == sc._norm(S.get(f"fund_{i}_kode") or nama):
                    mi = rd["mi"]
        e = fr.generic_entry(nama, mi)
    return e


def on_ffs_fetch(i: int):
    e = _fund_entry(i)
    extra = [u.strip() for u in str(S.get(f"fund_{i}_url", "")).split() if u.strip()]
    y, m = int(S.f_tahun), sc.BULAN.index(S.f_bulan)
    try:
        res = fr.fetch_latest_ffs(e, y, m, extra_urls=extra)
    except Exception as ex:
        flash("error", f"Gagal mengambil FFS: {ex}")
        return
    S[f"fund_{i}_log"] = res.log
    if res.ok and res.alokasi:
        a = res.alokasi
        S[f"fund_{i}_saham"], S[f"fund_{i}_obl"], S[f"fund_{i}_pu"] = a.saham, a.obligasi, a.pasar_uang
        tag = f"{sc.BULAN[a.as_of[1]]} {a.as_of[0]}" if a.as_of else "tanggal ?"
        S[f"fund_{i}_sumber"] = f"FFS {tag} — {res.url}"
        if extra:  # simpan link tambahan ke katalog
            cat = S.catalog
            ent = fr.find_in_catalog(cat, e["nama"])
            if ent is None:
                ent = dict(e)
                cat.append(ent)
            for u in extra:
                if u not in ent.setdefault("urls", []):
                    ent["urls"].insert(0, u)
            fr.save_catalog(cat)
        if a.as_of and a.as_of != (y, m):
            flash("warning", f"FFS '{e['nama']}' yang ditemukan adalah data {tag}, bukan {sc.BULAN[m]} {y}. "
                             "Periksa apakah FFS bulan laporan sudah terbit.")
        else:
            flash("success", f"FFS '{e['nama']}' ({tag}) terbaca: Saham {fmt_id(a.saham)}%, "
                             f"Obligasi {fmt_id(a.obligasi)}%, Pasar Uang/Lainnya {fmt_id(a.pasar_uang)}%.")
    else:
        flash("error", f"FFS '{e['nama']}' tidak berhasil diambil otomatis. Buka link FFS di bawah, unduh "
                       "PDF-nya, lalu upload di kotak 'Upload PDF FFS' atau isi manual.")


def on_ffs_pdf(i: int):
    up = S.get(f"fund_{i}_pdf")
    if up is None:
        return
    try:
        a, text = fr.parse_ffs_pdf(up.getvalue())
    except Exception as ex:
        flash("error", f"PDF tidak bisa dibaca: {ex}")
        return
    if a is None:
        S[f"fund_{i}_pdftext"] = text[:3000]
        flash("error", "Alokasi aset tidak ditemukan di PDF. Isi persentase secara manual "
                       "(teks PDF ditampilkan di bawah untuk dicek).")
        return
    S[f"fund_{i}_saham"], S[f"fund_{i}_obl"], S[f"fund_{i}_pu"] = a.saham, a.obligasi, a.pasar_uang
    tag = f"{sc.BULAN[a.as_of[1]]} {a.as_of[0]}" if a.as_of else "tanggal ?"
    S[f"fund_{i}_sumber"] = f"FFS {tag} — file {up.name}"
    flash("success", f"PDF FFS terbaca ({tag}): Saham {fmt_id(a.saham)}%, Obligasi {fmt_id(a.obligasi)}%, "
                     f"Pasar Uang/Lainnya {fmt_id(a.pasar_uang)}%.")


def on_add_fund():
    funds = [asdict(f) for f in collect_inputs().funds]
    if len(funds) >= 6:
        return
    funds.append(asdict(sc.FundSlot()))
    set_funds(funds)
    S[f"fund_{len(funds) - 1}_pilih"] = OTHER_FUND


def on_del_fund(i: int):
    funds = [asdict(f) for f in collect_inputs().funds]
    if len(funds) <= 1:
        return
    funds.pop(i)
    set_funds(funds)


def on_reset_template():
    S.raw = None
    S.result = None
    init_state(force_template=True)
    flash("info", "Semua input dikembalikan ke nilai template.")


def on_template_upload():
    up = S.get("up_template")
    if up is None:
        return
    p = sc.BASE_DIR / "template" / "Template_Pengguna.xlsx"
    p.write_bytes(up.getvalue())
    try:
        sc.template_raw_headers(p)
        sc.defaults_from_template(p)
    except Exception as e:
        flash("error", f"Template tidak cocok dengan format Stress Test: {e}")
        return
    S.template_path = str(p)
    flash("success", "Template diganti. Nilai input tetap; tekan 'Reset ke nilai template' bila ingin "
                     "memakai angka dari template baru.")


# --------------------------------------------------------------------------
# Sidebar
# --------------------------------------------------------------------------
with st.sidebar:
    logo = sc.BASE_DIR / "assets" / "logo_dpbca.png"
    if logo.exists():
        st.image(str(logo), width=170)
    st.markdown("### Stress Test Risiko")
    st.caption("Risiko Pasar · Kredit · Likuiditas (Permintaan OJK)")
    c1, c2 = st.columns([3, 2])
    c1.selectbox("Bulan laporan", sc.BULAN[1:], key="f_bulan")
    c2.number_input("Tahun", min_value=2020, max_value=2100, step=1, key="f_tahun")
    st.markdown(f"**File hasil:**  \n`{sc.output_filename(int(S.f_tahun), sc.BULAN.index(S.f_bulan))}`")
    status_box = st.container()
    st.divider()
    with st.expander("Template & pengaturan"):
        st.caption(f"Template aktif: `{sc.Path(S.template_path).name}`. Nilai awal input diambil dari "
                   f"{S.get('defaults_source', 'template')}.")
        st.file_uploader("Ganti template (file Stress Test bulan lalu, format sama)", type=["xlsx"],
                         key="up_template", on_change=on_template_upload)
        _w(st.button, "Reset ke nilai template", on_click=on_reset_template)

# --------------------------------------------------------------------------
# Header & pesan
# --------------------------------------------------------------------------
st.markdown(
    '<div class="dp-head"><div><div class="t1">Stress Test Risiko — Permintaan Data OJK</div>'
    '<div class="t2">Isi tab 1 → 7 berurutan. Semua angka bisa diubah manual; tombol otomatis hanya membantu.'
    '</div></div></div>', unsafe_allow_html=True)
# Wadah pesan selalu ada (1 elemen) agar posisi tab tidak bergeser -> tab aktif tidak kembali ke tab 1
with st.container():
    for kind, text in S.get("msgs", []):
        getattr(st, kind)(text)
S.msgs = []

tabs = st.tabs(["1 · RAW Investasi", "2 · IHSG, Yield & Kurs", "3 · Kalkulator Baseline",
                "4 · Komposisi RD (FFS)", "5 · Risiko Kredit", "6 · Risiko Likuiditas",
                "7 · Ringkasan & Buat Excel"])

# ==========================================================================
# TAB 1 — RAW INVESTASI
# ==========================================================================
with tabs[0]:
    st.subheader("Data RAW Investasi")
    st.markdown("Upload file **Rekap Investasi** (unduhan DISK). Aplikasi mencari sendiri tabelnya "
                "(kolom `JenisInvestasi`, `NilaiWajar`, dst.) lalu menyalin seluruhnya ke sheet "
                "**RAW INVESTASI** dengan urutan kolom template.")
    st.file_uploader("File Rekap Investasi (.xlsx)", type=["xlsx", "xlsm"], key="up_raw",
                     on_change=on_raw_upload)
    raw = S.raw
    if raw is None:
        st.info("Belum ada data. Setelah file di-upload, data kredit (Non-MI & MI), Nilai Wajar reksa dana, "
                "dan SPI saham terisi otomatis.")
    else:
        smy = sc.raw_summary(raw)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Jumlah baris", f"{smy['n_rows']:,}".replace(",", "."))
        m2.metric("Periode di file", f"{sc.BULAN[raw.period.month]} {raw.period.year}" if raw.period else "–")
        m3.metric("Total Nilai Wajar", f"Rp {fmt_id(smy['total_nw'] / 1e9, 2)} M")
        m4.metric("Sheet sumber", raw.sheet_name)
        if raw.period and (raw.period.year, raw.period.month) != (int(S.f_tahun), sc.BULAN.index(S.f_bulan)):
            st.warning(f"Periode file ({sc.BULAN[raw.period.month]} {raw.period.year}) berbeda dengan "
                       f"periode yang dipilih di samping ({S.f_bulan} {S.f_tahun}).")
        for w in raw.warnings:
            st.error(w)
        if raw.missing:
            st.warning("Kolom template yang tidak ada di file (dibiarkan kosong): " + ", ".join(raw.missing))
        if raw.extra:
            st.caption("Kolom tambahan di file yang tidak dipakai: " + ", ".join(map(str, raw.extra)))
        if smy["unknown_jenis"]:
            st.warning("Kode JenisInvestasi baru yang belum dikenal: " + ", ".join(smy["unknown_jenis"])
                       + ". Periksa apakah perlu masuk perhitungan.")
        cA, cB = st.columns([1, 1])
        with cA:
            st.markdown("**Nilai Wajar per JenisInvestasi** (sama dengan sheet CHECK)")
            df = pd.DataFrame([(k, v / 1e9) for k, v in smy["by_jenis"].items()],
                              columns=["JenisInvestasi", "Nilai Wajar (miliar)"])
            _w(st.dataframe, df.style.format({"Nilai Wajar (miliar)": lambda x: fmt_id(x, 2)}),
                         hide_index=True)
        with cB:
            st.markdown("**Reksa dana di portofolio**")
            rdl = pd.DataFrame(sc.reksa_dana_list(raw))
            if not rdl.empty:
                rdl = rdl.rename(columns={"nama": "Reksa Dana", "mi": "Manajer Investasi", "jenis": "Jenis",
                                          "np": "Nilai Perolehan", "nw": "Nilai Wajar"})
                _w(st.dataframe, rdl.style.format({"Nilai Perolehan": lambda x: fmt_id(x, 0),
                                               "Nilai Wajar": lambda x: fmt_id(x, 0)}),
                             hide_index=True)
        with st.expander("Lihat 50 baris pertama RAW"):
            _w(st.dataframe, pd.DataFrame(raw.rows[:50], columns=raw.headers))
        st.button("↻ Perbarui data turunan dari RAW (kredit, MI, Nilai Wajar)", on_click=refresh_from_raw)

# ==========================================================================
# TAB 2 — IHSG, YIELD, KURS
# ==========================================================================
inp_now = collect_inputs()
base = sc.compute_baseline(inp_now)
with tabs[1]:
    st.subheader("Skenario IHSG, Yield SBN 10 Tahun, dan Kurs USD/IDR")
    st.caption("Existing = kondisi akhir bulan laporan · Baseline = proyeksi 31 Desember (dihitung di tab 3) · "
               "Escalated = adjustment tim MR · Severe = ketentuan OJK.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### IHSG")
        st.number_input("IHSG Existing (penutupan akhir bulan)", min_value=0.0, step=1.0, format="%.2f",
                        key="f_ihsg_existing")
        st.number_input("IHSG Severe (ketentuan OJK)", min_value=0.0, step=1.0, format="%.2f",
                        key="f_ihsg_severe")
        ib = base["ihsg_baseline"]
        st.markdown(f'<div class="dp-card">IHSG Baseline (otomatis): <b>{fmt_id(ib)}</b></div>',
                    unsafe_allow_html=True)
        lo, hi = sorted([float(S.f_ihsg_severe), float(ib)])
        smin, smax = float(math.floor(lo) + 1), float(math.ceil(hi) - 1)
        if smax - smin >= 1:
            cur = float(S.f_ihsg_escalated)
            if not (smin <= cur <= smax):
                S.f_ihsg_escalated = float(round((smin + smax) / 2))
            elif cur != round(cur):
                S.f_ihsg_escalated = float(round(cur))
            st.slider("IHSG Escalated (tentukan sendiri, di antara Severe dan Baseline)",
                      min_value=smin, max_value=smax, step=1.0, key="f_ihsg_escalated", format="%.0f")
            st.caption(f"Rentang yang diizinkan: {fmt_id(smin, 0)} – {fmt_id(smax, 0)} "
                       "(geser, atau klik slider lalu pakai tombol panah ← → untuk per poin).")
        else:
            st.warning("IHSG Severe harus lebih rendah dari IHSG Baseline.")
            st.number_input("IHSG Escalated", step=1.0, format="%.2f", key="f_ihsg_escalated")
        ie, isv = float(S.f_ihsg_escalated), float(S.f_ihsg_severe)
        st.caption(f"% perubahan dari Baseline → Escalated **{pct_id((ie - ib) / ib if ib else None)}**, "
                   f"Severe **{pct_id((isv - ib) / ib if ib else None)}**")
        st.link_button("IHSG historis (Investing.com)", "https://id.investing.com/indices/idx-composite-historical-data")
    with c2:
        st.markdown("#### Yield SBN 10 Tahun (%)")
        st.number_input("Yield Existing (%)", step=0.001, format="%.4f", key="f_yield_existing")
        yb = base["yield_baseline"]
        st.markdown(f'<div class="dp-card">Yield Awal (A) = Yield Baseline (otomatis, hasil "goal seek"): '
                    f'<b>{pct_id(yb, 4)}</b></div>', unsafe_allow_html=True)
        for key, nm, yfield in (("esc", "Escalated", "f_yield_escalated"), ("sev", "Severe", "f_yield_severe")):
            manual_b = bool(S.get(f"f_b_manual_{key}"))
            if manual_b:   # Yield skenario mengikuti Yield awal + B
                S[yfield] = round(base[f"yield_{'escalated' if key == 'esc' else 'severe'}"] * 100, 8)
            else:          # B mengikuti Yield skenario - Yield awal
                S[f"f_b_{key}"] = round(base[f"shift_{'escalated' if key == 'esc' else 'severe'}"] * 100, 8)
            with st.container(border=True):
                st.markdown(f"**{nm}**")
                st.checkbox("Isi Peningkatan Suku Bunga (B) manual", key=f"f_b_manual_{key}")
                if S.get(f"f_b_manual_{key}"):
                    st.number_input(f"Peningkatan Suku Bunga (B) {nm} (%)", step=0.01, format="%.4f",
                                    key=f"f_b_{key}")
                    y_sk = yb + float(S[f"f_b_{key}"]) / 100
                    st.caption(f"Yield {nm} = Yield Awal {pct_id(yb, 4)} + B {fmt_id(float(S[f'f_b_{key}']), 4)}% "
                               f"= **{pct_id(y_sk, 4)}**")
                else:
                    st.number_input(f"Yield {nm} (%)", step=0.01, format="%.4f", key=yfield)
                    st.caption(f"Peningkatan Suku Bunga (B) = Yield {nm} − Yield Awal, dibulatkan = "
                               f"**{pct_id(base['shift_' + ('escalated' if key == 'esc' else 'severe')])}**")
        st.caption("B dipakai di sheet 1-Risiko Pasar (baris Skenario kenaikan suku bunga) dan "
                   "3-Risiko Likuiditas (Peningkatan suku bunga).")
        st.link_button("Yield SBN 10Y historis (Investing.com)",
                       "https://id.investing.com/rates-bonds/indonesia-10-year-bond-yield-historical-data")
    with c3:
        st.markdown("#### Kurs USD/IDR")
        st.number_input("Kurs Existing", step=1.0, format="%.2f", key="f_kurs_existing")
        st.number_input("Kurs Baseline (proyeksi analisis BCA)", step=1.0, format="%.2f", key="f_kurs_baseline")
        st.number_input("Kurs Escalated", step=1.0, format="%.2f", key="f_kurs_escalated")
        st.number_input("Kurs Severe", step=1.0, format="%.2f", key="f_kurs_severe")

# ==========================================================================
# TAB 3 — KALKULATOR BASELINE & PROYEKSI
# ==========================================================================
with tabs[2]:
    st.subheader("Kalkulator Skenario Baseline & Proyeksi Desember")
    with st.container(border=True):
        st.markdown("**Isi otomatis dari Budget & Realisasi (opsional)** — semua kolom di bawah tetap bisa "
                    "diketik manual.")
        cu1, cu2 = st.columns([3, 1])
        cu1.file_uploader("File Budget & Realisasi (.xlsx)", type=["xlsx", "xlsm"], key="up_budget")
        cu2.write("")
        _w(cu2.button, "Isi otomatis", on_click=on_budget_fill, type="primary")
        bd = S.get("budget")
        if bd is not None:
            lab = {"ex_total_aset": "Existing · Total Aset Investasi (Rata-rata NW)", "ex_spi": "Existing · SPI",
                   "ex_laba": "Existing · Laba Usaha (PIB YTD)", "ex_delta_spi": "Existing · Delta SPI YTD",
                   "pr_total_aset": "Proyeksi Des · Total Aset Investasi", "pr_spi": "Proyeksi Des · SPI",
                   "pr_laba": "Proyeksi Des · Laba Usaha", "pr_delta_spi": "Proyeksi Des · Delta SPI",
                   "spi_saham_des_juta": "SPI Saham 31 Des (juta)", "spi_saham_rep_juta": "SPI Saham bulan laporan (juta)",
                   "yc_delta_spi": "Delta SPI selain saham (miliar)", "aset_neto": "Aset Neto (Rp)"}
            rows = [(lab[k], fmt_id(v, 4 if abs(v) < 10 else 2), bd.sources.get(k, ""))
                    for k, v in bd.values.items() if k in lab]
            with st.expander(f"Sumber angka dari Budget ({len(rows)} angka)"):
                _w(st.dataframe, pd.DataFrame(rows, columns=["Isian", "Nilai", "Sel sumber"]), hide_index=True)
            for w in bd.warnings:
                st.warning(w)

    k1, k2 = st.columns(2)
    with k1:
        with st.container(border=True):
            st.markdown("#### Calculator IHSG Skenario Baseline")
            st.number_input("SPI Saham proyeksi 31 Des (juta Rp) — Budget, ASUMSI HU › SPI › Saham",
                            step=1.0, format="%.2f", key="f_spi_saham_des_juta")
            st.number_input("SPI Saham bulan laporan (juta Rp) — Budget / RAW", step=1.0, format="%.2f",
                            key="f_spi_saham_rep_juta")
            st.caption(f"SPI Reksa Dana Saham dari tab 4: {fmt_id(base['rd_spi_saham_juta'], 2)} juta")
            st.markdown(
                f"| | |\n|---|---|\n"
                f"| SPI 31 Des (L8) | =({int(base['L8a'])})+({int(base['L8b'])}) = **{fmt_id(base['L8'], 0)}** |\n"
                f"| SPI Reporting (L9) | **{fmt_id(base['L9'], 0)}** |\n"
                f"| % Delta SPI | {pct_id(base['delta_spi_pct'], 4)} |\n"
                f"| IHSG Existing | {fmt_id(inp_now.ihsg_existing)} |\n"
                f"| **IHSG Skenario Baseline** | **{fmt_id(base['ihsg_baseline'])}** |")
    with k2:
        with st.container(border=True):
            st.markdown("#### Calculator Yield Skenario Baseline")
            st.number_input("Nilai Wajar (miliar) — total instrumen terdampak suku bunga", step=1.0,
                            format="%.2f", key="f_yc_nilai_wajar")
            if S.raw is not None:
                sug = sc.total_terdampak_suku_bunga(S.raw, inp_now)
                st.caption(f"Saran dari RAW (= sheet 1-Risiko Pasar D15): {fmt_id(sug, 2)} → "
                           f"{fmt_id(sc.round_half_up(sug), 0)}")
            st.number_input("Delta SPI selain saham (miliar) — Des dikurangi bulan laporan", step=0.001,
                            format="%.3f", key="f_yc_delta_spi")
            st.number_input("Tenor (tahun)", min_value=1, step=1, key="f_yc_tenor")
            st.markdown(
                f"| | |\n|---|---|\n"
                f"| Bunga Existing | {pct_id(inp_now.yield_existing, 4)} |\n"
                f"| PV Existing (D5) | {fmt_id(base['pv_existing'], 8)} |\n"
                f"| PV Skenario (C5, hasil goal seek) | {fmt_id(base['pv_skenario'], 8)} |\n"
                f"| **Yield Skenario Baseline** | **{pct_id(base['yield_baseline'], 4)}** |")
            st.caption("Goal seek di Excel dihitung langsung: C5 = Delta SPI / Nilai Wajar + PV Existing.")

    p1, p2 = st.columns(2)
    with p1:
        with st.container(border=True):
            st.markdown("#### Existing — rata-rata investasi (miliar)")
            st.number_input("Total Aset Investasi", step=0.01, format="%.6f", key="f_ex_total_aset")
            st.number_input("SPI", step=0.01, format="%.6f", key="f_ex_spi")
            st.number_input("Laba Usaha", step=0.01, format="%.6f", key="f_ex_laba")
            st.number_input("Delta SPI", step=0.01, format="%.6f", key="f_ex_delta_spi")
    with p2:
        with st.container(border=True):
            st.markdown(f"#### Proyeksi Desember {S.f_tahun} (miliar)")
            st.number_input("Total Aset Investasi (Rata-rata Nilai Wajar)", step=0.01, format="%.6f",
                            key="f_pr_total_aset")
            st.number_input("SPI ", step=0.01, format="%.6f", key="f_pr_spi")
            st.number_input("Laba Usaha (Pendapatan Investasi Bersih YTD)", step=0.01, format="%.6f",
                            key="f_pr_laba")
            st.number_input("Delta SPI (YTD)", step=0.01, format="%.6f", key="f_pr_delta_spi")
    st.number_input("Aset Neto (Rp) — LAN Aset Neto OJK DISK", step=1.0, format="%.0f", key="f_aset_neto")
    st.caption(f"= Rp {fmt_id(S.f_aset_neto, 0)}")

# ==========================================================================
# TAB 4 — KOMPOSISI RD
# ==========================================================================
with tabs[3]:
    st.subheader("Komposisi Reksa Dana Campuran (Fund Fact Sheet)")
    st.markdown("Default: **Schroder Dana Terpadu II** dan **Trimegah Balanced Absolute Strategy Kelas A**. "
                "Bila diganti, pilih reksa dana lain dari daftar (katalog + reksa dana di RAW) atau ketik namanya. "
                "Tombol **Ambil FFS terbaru** mengunduh fund fact sheet bulan laporan dan membaca porsi "
                "Saham / Obligasi / Pasar Uang.")
    cat_names = [f["nama"] for f in S.catalog]
    raw_names = []
    if S.raw is not None:
        for rd in sc.reksa_dana_list(S.raw):
            if all(sc._norm(rd["nama"]) != sc._norm(x) for x in cat_names + raw_names):
                raw_names.append(rd["nama"])
    y_sel, m_sel = int(S.f_tahun), sc.BULAN.index(S.f_bulan)
    for i in range(S.n_funds):
        with st.container(border=True):
            cur = S.get(f"fund_{i}_pilih", "")
            options = list(dict.fromkeys(cat_names + raw_names + ([cur] if cur and cur != OTHER_FUND else [])
                                         + [OTHER_FUND]))
            if cur not in options:
                S[f"fund_{i}_pilih"] = options[0]
            h1, h2 = st.columns([5, 1])
            h1.selectbox(f"Reksa dana #{i + 1}", options, key=f"fund_{i}_pilih", on_change=on_fund_select,
                         args=(i,))
            h2.write("")
            _w(h2.button, "Hapus", key=f"del_{i}", on_click=on_del_fund, args=(i,), disabled=S.n_funds <= 1)
            if S.get(f"fund_{i}_pilih") == OTHER_FUND:
                st.text_input("Nama reksa dana (tertulis di sheet Komposisi RD)", key=f"fund_{i}_nama")
            a1, a2, a3, a4 = st.columns(4)
            a1.number_input("Saham (%)", step=0.01, format="%.2f", key=f"fund_{i}_saham")
            a2.number_input("Obligasi (%)", step=0.01, format="%.2f", key=f"fund_{i}_obl")
            a3.number_input("Pasar Uang / Lainnya (%)", step=0.01, format="%.2f", key=f"fund_{i}_pu")
            a4.number_input("Nilai Wajar (Rp) — dari RAW", step=1.0, format="%.2f", key=f"fund_{i}_nw")
            tot = S[f"fund_{i}_saham"] + S[f"fund_{i}_obl"] + S[f"fund_{i}_pu"]
            ok = abs(tot - 100) <= 1.0
            st.markdown(f"Total komposisi: <span class='{'dp-ok' if ok else 'dp-bad'}'>{fmt_id(tot)}%</span> · "
                        f"Di Excel: B = {fmt_id(S[f'fund_{i}_saham'])}% · C = "
                        f"`={sc.fmt_pct_literal(S[f'fund_{i}_obl'])}%+{sc.fmt_pct_literal(S[f'fund_{i}_pu'])}%`",
                        unsafe_allow_html=True)
            if S.get(f"fund_{i}_sumber"):
                st.caption("Sumber: " + S[f"fund_{i}_sumber"])
            b1, b2 = st.columns([1, 2])
            with b1:
                _w(st.button, "🔄 Ambil FFS terbaru", key=f"fetch_{i}", on_click=on_ffs_fetch, args=(i,),
                          type="primary")
                st.file_uploader("…atau upload PDF FFS", type=["pdf"], key=f"fund_{i}_pdf",
                                 on_change=on_ffs_pdf, args=(i,))
            with b2:
                e = _fund_entry(i)
                with st.expander("Link fund fact sheet"):
                    for u in e.get("urls", []):
                        st.markdown(f"- [{fr.expand_url(u, y_sel, m_sel)}]({fr.expand_url(u, y_sel, m_sel)})")
                    for u in e.get("halaman", []):
                        st.markdown(f"- Halaman produk: [{u}]({u})")
                    for lab, u in fr.search_links(S.get(f"fund_{i}_nama") or "", y_sel, m_sel):
                        st.markdown(f"- [{lab}]({u})")
                    st.text_input("Tambah link FFS sendiri (dipakai & disimpan saat 'Ambil FFS terbaru')",
                                  key=f"fund_{i}_url", placeholder="https://.../factsheet.pdf")
                if S.get(f"fund_{i}_log"):
                    with st.expander("Log pengambilan FFS"):
                        st.code("\n".join(S[f"fund_{i}_log"]))
                if S.get(f"fund_{i}_pdftext"):
                    with st.expander("Teks PDF (untuk cek manual)"):
                        st.text(S[f"fund_{i}_pdftext"])
    st.button("➕ Tambah reksa dana", on_click=on_add_fund, disabled=S.n_funds >= 6)

    st.markdown("#### SPI Reksa Dana per Jenis (input manual dari DISK: *Laporan SPI Reksadana per Jenis*)")
    r1, r2 = st.columns(2)
    with r1:
        with st.container(border=True):
            st.markdown("**SAHAM**")
            st.number_input("Nilai Perolehan (Rp)", step=1.0, format="%.2f", key="f_rd_saham_np")
            st.number_input("Nilai Wajar (Rp)", step=1.0, format="%.2f", key="f_rd_saham_nw")
            st.caption(f"SPI = Rp {fmt_id(S.f_rd_saham_nw - S.f_rd_saham_np)}")
    with r2:
        with st.container(border=True):
            st.markdown("**CAMPURAN**")
            st.number_input("Nilai Perolehan (Rp) ", step=1.0, format="%.2f", key="f_rd_campuran_np")
            st.number_input("Nilai Wajar (Rp) ", step=1.0, format="%.2f", key="f_rd_campuran_nw")
            st.caption(f"SPI = Rp {fmt_id(S.f_rd_campuran_nw - S.f_rd_campuran_np)}")
    st.button("Isi dari RAW (lalu cocokkan dengan laporan DISK)", on_click=on_rd_from_raw)

# ==========================================================================
# TAB 5 — RISIKO KREDIT
# ==========================================================================
with tabs[4]:
    st.subheader("Risiko Kredit")
    st.markdown("Eksposur dihitung dari RAW (Nilai Wajar per grup/pihak; tidak termasuk SBI & SBN, "
                "termasuk DOC, Deposito, Obligasi & Sukuk korporasi). **Rating grup** diambil dari daftar "
                "rating tersimpan — isi/ubah bila ada grup baru atau rating berubah (sumber: Display "
                "Monitoring Risiko Investasi).")
    pd1, pd2, pd3, pd4, pd5 = st.columns(5)
    for col, kel in zip((pd1, pd2, pd3, pd4), ("AAA", "AA", "A", "D")):
        S.pd_kelas[kel] = col.number_input(f"PD {sc.KELAS_LABEL[kel]} (%)",
                                           value=float(S.pd_kelas.get(kel, sc.DEFAULT_PD_KELAS[kel]) * 100),
                                           step=1.0, key=f"pd_{kel}") / 100
    pd5.number_input("LGD (%)", step=1.0, key="f_lgd")
    st.markdown("**Non Manajer Investasi** (sheet 2-Risiko Kredit)")
    edited = _w(st.data_editor, 
        S.credit_df, key=f"cred_ed_{S.credit_ver}", num_rows="dynamic",
        hide_index=True,
        column_config={
            "Grup / Pihak": st.column_config.TextColumn(width="large"),
            "Eksposur (miliar)": st.column_config.NumberColumn(format="%.6f"),
            "Rating": st.column_config.TextColumn(help="Contoh: idAAA, AAA(idn), idAA-, idA+, idD"),
            "Kelas": st.column_config.SelectboxColumn(options=KELAS_OPS,
                                                      help="Kosongkan = otomatis dari rating"),
        })
    S.credit_edited = edited
    inp_now = collect_inputs()
    prev = pd.DataFrame([{"No": i + 1, "Grup": c.nama, "Eksposur": c.eksposur, "Rating": c.rating,
                          "Kelas": c.kelas or "⚠ belum", "PD": c.pd, "Hasil (PD×LGD×EAD)": c.pd * inp_now.lgd * c.eksposur}
                         for i, c in enumerate(inp_now.credit)])
    missing = [c.nama for c in inp_now.credit if not c.rating or c.kelas not in sc.DEFAULT_PD_KELAS]
    if missing:
        st.error("Rating/kelas belum lengkap: " + ", ".join(missing))
    nonmi = float(prev["Hasil (PD×LGD×EAD)"].sum()) if not prev.empty else 0.0
    st.markdown("**Manajer Investasi** (sheet 2-Check (2))")
    medited = _w(st.data_editor, 
        S.mi_df, key=f"mi_ed_{S.mi_ver}", num_rows="dynamic", hide_index=True,
        column_config={
            "Pihak (RAW)": st.column_config.TextColumn(disabled=True),
            "Nominal (miliar)": st.column_config.NumberColumn(format="%.6f"),
            "Rating": st.column_config.SelectboxColumn(options=MI_OPS, required=False),
        })
    S.mi_edited = medited
    cm = st.columns(len(sc.DEFAULT_PD_MI) + 1)
    for col, k in zip(cm, sc.DEFAULT_PD_MI):
        S.pd_mi[k] = col.number_input(f"PD MI '{k}' (%)", value=float(S.pd_mi.get(k, sc.DEFAULT_PD_MI[k]) * 100),
                                      step=1.0, key=f"pdmi_{k}") / 100
    inp_now = collect_inputs()
    mi_res = sum(inp_now.pd_mi.get(m.rating, 0.0) * inp_now.lgd * m.nominal for m in inp_now.mi)
    bad_mi = [m.nama for m in inp_now.mi if m.rating not in inp_now.pd_mi]
    if bad_mi:
        st.error("Rating MI belum dipilih: " + ", ".join(bad_mi))
    x1, x2, x3, x4 = st.columns(4)
    x1.metric("Hasil kredit Non-MI", fmt_id(nonmi, 4))
    x2.metric("Hasil kredit MI", fmt_id(mi_res, 4))
    x3.metric("Total Hasil Stress Test Kredit (miliar)", fmt_id(nonmi + mi_res, 4))
    x4.metric("% dari Aset Neto", pct_id((nonmi + mi_res) * 1e9 / inp_now.aset_neto if inp_now.aset_neto else None, 3))
    with st.expander("Rincian perhitungan Non-MI"):
        if not prev.empty:
            _w(st.dataframe, prev.style.format({"Eksposur": lambda x: fmt_id(x, 6), "PD": lambda x: pct_id(x, 0),
                                            "Hasil (PD×LGD×EAD)": lambda x: fmt_id(x, 4)}),
                         hide_index=True)
    st.button("↻ Muat ulang eksposur dari RAW", on_click=refresh_from_raw, disabled=S.raw is None)

# ==========================================================================
# TAB 6 — RISIKO LIKUIDITAS
# ==========================================================================
with tabs[5]:
    st.subheader("Risiko Likuiditas")
    l1, l2 = st.columns(2)
    with l1:
        st.number_input("Eksposur risiko likuiditas — Total 1 dan 2 (Rp): keterlambatan iuran + "
                        "peningkatan pembayaran manfaat pensiun", step=1_000_000.0, format="%.0f",
                        key="f_likuiditas_eksposur")
        st.caption(f"= Rp {fmt_id(S.f_likuiditas_eksposur, 0)}")
        st.number_input("Asumsi tenor surat berharga yang dijual (tahun)", min_value=1, step=1,
                        key="f_likuiditas_tenor")
    with l2:
        st.markdown("**Parameter per skenario** — (a) Peningkatan suku bunga dari tab 2, "
                    "(b) Yield awal = Yield Baseline.")
        lp = sc.liquidity_params(collect_inputs())
        rows = []
        for key, nm in (("esc", "Escalated"), ("sev", "Severe")):
            q = lp[key]
            rows.append({"Skenario": nm,
                         "Kenaikan suku bunga (a)": pct_id(q["a"], 2 if abs(q["a"] * 1e4 - round(q["a"] * 1e4)) < 1e-6 else 4), "Yield awal (b)": pct_id(q["b"], 4),
                         "Yield skenario (c)": pct_id(q["c"], 4),
                         "Hasil (miliar)": fmt_id(q["hasil"] / 1e9, 4), "% Aset Neto": pct_id(q["pct_aset"], 4)})
        _w(st.dataframe, pd.DataFrame(rows), hide_index=True)

# ==========================================================================
# TAB 7 — RINGKASAN & BUAT EXCEL
# ==========================================================================
inp_final = collect_inputs()
errors, warns = sc.validate_inputs(inp_final, S.raw)
with tabs[6]:
    st.subheader("Ringkasan & Buat Excel")
    y_, m_ = inp_final.tahun, inp_final.bulan
    st.markdown(f"Periode: **{sc.BULAN[m_]} {y_}** (posisi {sc.month_end(y_, m_).strftime('%d-%m-%Y')}) → "
                f"file **{sc.output_filename(y_, m_)}**")
    if errors:
        st.error("Belum bisa membuat Excel:\n\n" + "\n".join(f"- {e}" for e in errors))
    for w in warns:
        st.warning(w)
    go = _w(st.button, "🧮 Hitung & Buat Excel", type="primary", disabled=bool(errors))
    if go:
        with st.spinner("Menyusun workbook dari template dan menghitung semua rumus…"):
            try:
                res = sc.build_workbook(collect_inputs(), S.raw, S.template_path)
                sc.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                out_path = sc.OUTPUT_DIR / res.filename
                try:
                    out_path.write_bytes(res.content)
                    saved = str(out_path)
                except PermissionError:
                    saved = None
                S.result = res
                S.result_path = saved
                sc.save_last_inputs(collect_inputs())
                S.master = sc.update_master_from_inputs(S.master, collect_inputs())
                sc.save_master(S.master)
                fr.save_catalog(S.catalog)
            except Exception as e:
                S.result = None
                st.exception(e)
    res = S.get("result")
    if res is not None:
        sm = res.summary
        st.success(f"Selesai: **{res.filename}**" + (f" — tersimpan di `{S.result_path}`" if S.get("result_path")
                                                     else " (file sedang terbuka di Excel? tutup lalu ulangi)"))
        _w(st.download_button, "⬇️ Unduh Excel", data=res.content, file_name=res.filename, type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        g = st.columns(5)
        for col, (lab, val, pc, nd) in zip(g, [
                ("Pasar Escalated", sm["pasar_escalated"], sm["pasar_pct_aset_esc"], 2),
                ("Pasar Severe", sm["pasar_severe"], sm["pasar_pct_aset_sev"], 2),
                ("Kredit", sm["kredit"], sm["kredit_pct_aset"], 2),
                ("Likuiditas Escalated", sm["likuiditas_escalated"], sm["lik_pct_esc"], 3),
                ("Likuiditas Severe", sm["likuiditas_severe"], sm["lik_pct_sev"], 3)]):
            col.metric(lab, fmt_id(val, 2))
            col.caption(f"{pct_id(pc, nd)} dari Aset Neto")
        st.caption("Angka dalam miliar Rupiah.")

        def _f(row):
            out = dict(row)
            for c in ("Existing", "Baseline", "Escalated", "Severe"):
                v = row[c]
                out[c] = pct_id(v, 2) if row["Keterangan"] == "ROI" else fmt_id(v, 2)
            return out
        tb = pd.DataFrame([_f(r) for r in sm["tabel"]])
        st.markdown("**Tabel Hasil Stress Testing** (sama dengan sheet *Hasil Stress Testing*)")
        _w(st.dataframe, tb, hide_index=True, height=680)
        for w in res.warnings:
            st.warning(w)

# --------------------------------------------------------------------------
# Status di sidebar
# --------------------------------------------------------------------------
with status_box:
    checks = [
        ("RAW Investasi", S.raw is not None),
        ("IHSG & Yield", bool(inp_final.ihsg_existing and inp_final.yield_existing)),
        ("IHSG Escalated valid", sc.compute_baseline(inp_final)["ihsg_escalated_valid"]),
        ("Proyeksi & Aset Neto", bool(inp_final.pr_total_aset and inp_final.aset_neto)),
        ("Komposisi RD", all(abs(f.saham + f.obligasi + f.pasar_uang - 100) <= 1 for f in inp_final.funds)),
        ("Rating kredit lengkap", all(c.kelas in sc.DEFAULT_PD_KELAS for c in inp_final.credit)
         and all(m.rating in inp_final.pd_mi for m in inp_final.mi)),
    ]
    st.markdown("**Kelengkapan**")
    for lab, ok in checks:
        st.markdown(f"{'✅' if ok else '⬜'} {lab}")
    if not errors:
        st.success("Siap dibuat (tab 7)")
