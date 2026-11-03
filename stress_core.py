"""
stress_core.py
==============
Logika inti aplikasi "Stress Test Risiko" (Permintaan OJK: data stress test
risiko pasar, kredit, likuiditas).

Isi modul:
  * read_rekap()          : membaca file "Rekap Investasi" -> sheet RAW INVESTASI
  * read_budget()         : (opsional) membaca file "Budget & Realisasi"
  * defaults_from_template: nilai awal input diambil dari template / bulan lalu
  * compute_baseline()    : kalkulator IHSG & Yield skenario baseline
  * build_workbook()      : membuat file Excel final dari template (format sama
                            persis), menghitung semua rumus, dan menyimpan nilai
                            hasil hitungnya di dalam file.
"""
from __future__ import annotations

import calendar
import datetime as dt
import html
import io
import json
import math
import re
import zipfile
import xml.etree.ElementTree as ET
from copy import copy
from dataclasses import asdict, dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.cell_style import StyleArray

from xl_engine import Engine, XLErr

# --------------------------------------------------------------------------
# Lokasi file
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template" / "Template_Stress_Test_OJK.xlsx"
CONFIG_DIR = BASE_DIR / "config"
OUTPUT_DIR = BASE_DIR / "hasil"
MASTER_PATH = CONFIG_DIR / "master_rating.json"
LAST_INPUTS_PATH = CONFIG_DIR / "input_terakhir.json"

BULAN = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
BULAN_SINGKAT = ["", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agust",
                 "Sept", "Okt", "Nov", "Des"]
MONTH_EN = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept",
            "Oct", "Nov", "Dec"]

# Nama sheet di template
SH_HASIL = "Hasil Stress Testing"
SH_SB = "Skenario Baseline"
SH_KRD = "Komposisi RD"
SH_RAW = "RAW INVESTASI"
SH_PASAR_ESC = "1-Risiko Pasar Escalated"
SH_PASAR_SEV = "1-Risiko Pasar Severe"
SH_PASAR_BASE = "1-Risiko Pasar Baseline"
SH_PASAR_EXIST = "1-Risiko Pasar Existing"
SH_KREDIT = "2-Risiko Kredit"
SH_CHK1 = "2-CHECK (1)"
SH_CHK2 = "2-Check (2)"
SH_LIK_ESC = "3-Risiko Likuiditas Escalated"
SH_LIK_SEV = "3-Risiko Likuiditas Severed"
SH_CHECK = "CHECK"
SH_RATING = "Rating"

KELAS_LABEL = {"AAA": "Triple (AAA)", "AA": "Double (AA)", "A": "Single (A)", "D": "D (default)"}
DEFAULT_PD_KELAS = {"AAA": 0.20, "AA": 0.25, "A": 0.30, "D": 1.00}
DEFAULT_PD_MI = {"Sangat Bagus": 0.20, "Bagus": 0.25, "Cukup": 0.30}
DEFAULT_LGD = 0.45


# --------------------------------------------------------------------------
# Utilitas
# --------------------------------------------------------------------------
def month_end(year: int, month: int) -> dt.date:
    return dt.date(year, month, calendar.monthrange(year, month)[1])


def output_filename(year: int, month: int) -> str:
    return f"Stress Test Risiko {BULAN[month]} {year}.xlsx"


def round_half_up(x: float, nd: int = 0) -> float:
    q = Decimal(1).scaleb(-nd)
    return float(Decimal(repr(float(x))).quantize(q, rounding=ROUND_HALF_UP))


def miliar_trunc(rupiah: float) -> float:
    """Rupiah -> miliar, dipotong (bukan dibulatkan) per Rp1.000 = 6 desimal.
    Sama dengan cara angka eksposur ditulis di template (mis. 31,890905)."""
    ribu = math.floor(abs(rupiah) / 1000 + 1e-7)
    return math.copysign(ribu / 1e6, rupiah) if rupiah else 0.0


def _norm(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower()) if s is not None else ""


def _num(v, default=0.0) -> float:
    if v is None or v == "":
        return default
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return default


def fmt_pct_literal(v: float) -> str:
    """46.53 -> '46.53', 18.50 -> '18.5', 5.0 -> '5' (untuk rumus =47.91%+5.55%)."""
    s = f"{round_half_up(v, 4):.4f}".rstrip("0").rstrip(".")
    return s if s not in ("", "-0") else "0"


def smart_title(name: str) -> str:
    """'PT TRIMEGAH ASSET MANAGEMENT' -> 'PT Trimegah Asset Management'."""
    keep_upper = {"PT", "TBK", "BNP", "BNI", "BRI", "BCA", "BTN", "UOB", "CIMB",
                  "OCBC", "HSBC", "DBS", "RHB", "KB", "MNC", "II", "III", "IV",
                  "AM", "MI", "RD", "BPD", "SMBC", "MUFG", "KSEI", "BUMN"}
    out = []
    for w in str(name).split():
        core = re.sub(r"[^A-Za-z]", "", w).upper()
        if core in keep_upper:
            out.append(w.upper())
        else:
            out.append(w[:1].upper() + w[1:].lower())
    return " ".join(out)


def rating_class(rating: str | None) -> str | None:
    """Kelompokkan rating: 'AAA' (triple), 'AA' (double), 'A' (single), 'D'."""
    if not rating:
        return None
    s = str(rating).upper().strip()
    s = re.sub(r"\(.*?\)", "", s)          # (idn), (sy), (gg), ...
    s = re.sub(r"^ID", "", s).strip()
    s = s.replace("+", "").replace("-", "").strip()
    if s == "AAA":
        return "AAA"
    if s == "AA":
        return "AA"
    if s == "A":
        return "A"
    if s in ("D", "SD", "RD"):
        return "D"
    return None


def rating_sort_key(lab: str):
    """Urutan seperti pivot Excel: AAA(idn), idA-, idA+, idAA, idAA-, idAA+, idAAA, idD."""
    s = str(lab)
    suffix = 0
    base = s
    if s.endswith("-"):
        suffix, base = 1, s[:-1]
    elif s.endswith("+"):
        suffix, base = 2, s[:-1]
    return (base.lower(), suffix)


# --------------------------------------------------------------------------
# Struktur input
# --------------------------------------------------------------------------
@dataclass
class FundSlot:
    nama: str = ""
    kode_raw: str = ""          # nama reksa dana di kolom 'Kode' RAW INVESTASI
    saham: float = 0.0          # persen, mis. 46.53
    obligasi: float = 0.0       # persen
    pasar_uang: float = 0.0     # persen (Pasar Uang / Kas / Lainnya)
    nilai_wajar: float = 0.0    # Rupiah
    sumber: str = ""            # URL / nama file FFS


@dataclass
class CreditRow:
    nama: str = ""
    eksposur: float = 0.0       # miliar Rp
    rating: str = ""
    kelas: str = ""             # AAA / AA / A / D
    pd: float = 0.0


@dataclass
class MIRow:
    pihak: str = ""             # kunci dari RAW (kolom Pihak)
    nama: str = ""              # nama tampilan
    nominal: float = 0.0        # miliar Rp
    rating: str = ""            # Sangat Bagus / Bagus / Cukup


@dataclass
class StressInputs:
    tahun: int = 2026
    bulan: int = 9
    # Skenario pasar
    ihsg_existing: float = 0.0
    ihsg_escalated: float = 0.0
    ihsg_severe: float = 0.0
    yield_existing: float = 0.0      # pecahan, 0.07106 = 7,106%
    yield_escalated: float = 0.0
    yield_severe: float = 0.0
    kurs_existing: float = 0.0
    kurs_baseline: float = 0.0
    kurs_escalated: float = 0.0
    kurs_severe: float = 0.0
    # Rata-rata investasi existing (miliar)
    ex_total_aset: float = 0.0
    ex_spi: float = 0.0
    ex_laba: float = 0.0
    ex_delta_spi: float = 0.0
    # Proyeksi Desember (miliar)
    pr_total_aset: float = 0.0
    pr_spi: float = 0.0
    pr_laba: float = 0.0
    pr_delta_spi: float = 0.0
    aset_neto: float = 0.0           # Rupiah
    # Kalkulator yield baseline
    yc_nilai_wajar: float = 0.0      # miliar
    yc_delta_spi: float = 0.0        # miliar
    yc_tenor: int = 10
    # Kalkulator IHSG baseline (juta Rp)
    spi_saham_des_juta: float = 0.0
    spi_saham_rep_juta: float = 0.0
    rd_spi_saham_juta: float = 0.0   # SPI Reksa Dana Saham (juta) - input manual dari PDF
                                     # "Laporan SPI Reksadana per Jenis" (DISK), mis. -3041
    # SPI reksa dana per jenis (Rp) - Laporan SPI Reksadana per Jenis (DISK)
    rd_saham_np: float = 0.0
    rd_saham_nw: float = 0.0
    rd_campuran_np: float = 0.0
    rd_campuran_nw: float = 0.0
    # Tabel-tabel
    funds: list = field(default_factory=list)
    credit: list = field(default_factory=list)
    mi: list = field(default_factory=list)
    lgd: float = DEFAULT_LGD
    pd_kelas: dict = field(default_factory=lambda: dict(DEFAULT_PD_KELAS))
    pd_mi: dict = field(default_factory=lambda: dict(DEFAULT_PD_MI))
    # Likuiditas
    likuiditas_eksposur: float = 56_000_000_000.0
    likuiditas_tenor: int = 10
    # Peningkatan suku bunga (B) diisi manual (tab 2). Bila aktif:
    #   Yield skenario = Yield awal (Yield Baseline) + B
    # Bila tidak aktif: B = Yield skenario - Yield Baseline (dibulatkan 2 desimal persen).
    b_manual_esc: bool = False
    b_esc: float = 0.0               # pecahan, 0.01 = 1,00%
    b_manual_sev: bool = False
    b_sev: float = 0.0

    # ---- serialisasi ----
    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)

    @classmethod
    def from_dict(cls, d: dict) -> "StressInputs":
        d = dict(d)
        funds = [FundSlot(**f) for f in d.pop("funds", [])]
        credit = [CreditRow(**c) for c in d.pop("credit", [])]
        mi = [MIRow(**m) for m in d.pop("mi", [])]
        known = {k: v for k, v in d.items() if k in cls.__dataclass_fields__}
        obj = cls(**known)
        obj.funds, obj.credit, obj.mi = funds, credit, mi
        return obj


def save_last_inputs(inp: StressInputs, path: Path = LAST_INPUTS_PATH):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(inp.to_json(), encoding="utf-8")


def load_last_inputs(path: Path = LAST_INPUTS_PATH) -> StressInputs | None:
    try:
        return StressInputs.from_dict(json.loads(path.read_text(encoding="utf-8")))
    except Exception:
        return None


# --------------------------------------------------------------------------
# Perhitungan skenario baseline (sama dengan sheet "Skenario Baseline")
# --------------------------------------------------------------------------
def compute_baseline(inp: StressInputs) -> dict:
    out = {}
    rd_spi_saham_juta = float(inp.rd_spi_saham_juta)
    l8a = round_half_up(inp.spi_saham_des_juta)
    l8b = round_half_up(rd_spi_saham_juta)
    L8 = l8a + l8b
    L9 = round_half_up(inp.spi_saham_rep_juta + rd_spi_saham_juta)
    L4 = (L8 - L9) / L9 if L9 else 0.0
    out.update(rd_spi_saham_juta=rd_spi_saham_juta, L8a=l8a, L8b=l8b, L8=L8, L9=L9,
               delta_spi_pct=L4, ihsg_baseline=inp.ihsg_existing * (1 + L4))
    tenor = inp.yc_tenor or 10
    D5 = 1 / (1 + inp.yield_existing) ** tenor
    C5 = (inp.yc_delta_spi / inp.yc_nilai_wajar if inp.yc_nilai_wajar else 0.0) + D5
    y_base = (1 / C5) ** (1 / tenor) - 1 if C5 > 0 else float("nan")
    out.update(pv_existing=D5, pv_skenario=C5, yield_baseline=y_base)
    for key, nm, y_input in (("esc", "escalated", inp.yield_escalated), ("sev", "severe", inp.yield_severe)):
        if getattr(inp, f"b_manual_{key}"):
            b = float(getattr(inp, f"b_{key}"))
            out[f"shift_{nm}"] = b
            out[f"yield_{nm}"] = y_base + b          # Yield skenario = Yield awal + B
        else:
            out[f"shift_{nm}"] = round_half_up(y_input - y_base, 4) if y_base == y_base else float("nan")
            out[f"yield_{nm}"] = y_input
    lo, hi = sorted([inp.ihsg_severe, out["ihsg_baseline"]])
    out["ihsg_escalated_valid"] = lo < inp.ihsg_escalated < hi
    out["ihsg_range"] = (lo, hi)
    return out


def liquidity_params(inp: StressInputs, base: dict | None = None) -> dict:
    """Parameter & hasil sheet 3-Risiko Likuiditas per skenario ('esc', 'sev')."""
    base = base or compute_baseline(inp)
    t = int(inp.likuiditas_tenor or 10)
    out = {}
    for key, a in (("esc", base["shift_escalated"]), ("sev", base["shift_severe"])):
        b = base["yield_baseline"]
        c = a + b
        try:
            hasil = ((1 / (1 + b) ** t) - (1 / (1 + c) ** t)) * inp.likuiditas_eksposur
        except (ZeroDivisionError, OverflowError, ValueError, TypeError):
            hasil = float("nan")
        out[key] = {"a": a, "b": b, "c": c, "hasil": hasil,
                    "pct_aset": hasil / inp.aset_neto if inp.aset_neto else None}
    return out


# --------------------------------------------------------------------------
# RAW INVESTASI (Rekap Investasi)
# --------------------------------------------------------------------------
@dataclass
class RawData:
    headers: list
    rows: list
    sheet_name: str = ""
    header_row: int = 1
    period: dt.date | None = None
    missing: list = field(default_factory=list)
    extra: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    source_name: str = ""

    def col(self, name: str) -> int:
        n = _norm(name)
        for i, h in enumerate(self.headers):
            if _norm(h) == n:
                return i
        raise KeyError(name)

    def values(self, name: str):
        i = self.col(name)
        return [r[i] for r in self.rows]

    def records(self):
        idx = {_norm(h): i for i, h in enumerate(self.headers)}
        for r in self.rows:
            yield {k: r[i] for k, i in idx.items()}


def template_raw_headers(template_path=TEMPLATE_PATH) -> list:
    wb = openpyxl.load_workbook(template_path, read_only=True)
    ws = wb[SH_RAW]
    hdr = [c for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    wb.close()
    while hdr and hdr[-1] is None:
        hdr.pop()
    return hdr


REQUIRED_RAW = ["JenisInvestasi", "NilaiWajar", "NilaiPerolehan", "jatuhTempo", "Type", "Pihak", "Kode"]


def read_rekap(src, template_headers=None, source_name="") -> RawData:
    """Baca file Rekap Investasi (xlsx). Cari otomatis sheet & baris header."""
    if template_headers is None:
        template_headers = template_raw_headers()
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    found = None
    for ws in wb.worksheets:
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), 1):
            norms = [_norm(v) for v in row]
            if "jenisinvestasi" in norms and "nilaiwajar" in norms:
                found = (ws, r_idx, list(row))
                break
        if found:
            break
    if not found:
        wb.close()
        raise ValueError("Tabel Rekap Investasi tidak ditemukan. Pastikan file berisi kolom "
                         "'JenisInvestasi' dan 'NilaiWajar' (hasil unduhan Rekap Investasi).")
    ws, hr, hdr = found
    src_idx = {}
    for i, h in enumerate(hdr):
        if h is not None and _norm(h) not in src_idx:
            src_idx[_norm(h)] = i
    mapping = [src_idx.get(_norm(h)) for h in template_headers]
    missing = [h for h, m in zip(template_headers, mapping) if m is None]
    tnorm = {_norm(h) for h in template_headers}
    extra = [h for h in hdr if h is not None and _norm(h) not in tnorm]
    rows = []
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        if row is None or all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue
        rows.append([row[i] if (i is not None and i < len(row)) else None for i in mapping])
    sheet_name = ws.title
    wb.close()

    raw = RawData(headers=list(template_headers), rows=rows, sheet_name=sheet_name,
                  header_row=hr, missing=missing, extra=extra, source_name=source_name)
    miss_req = [m for m in REQUIRED_RAW if _norm(m) in {_norm(x) for x in missing}]
    if miss_req:
        raw.warnings.append("Kolom wajib tidak ditemukan: " + ", ".join(miss_req))
    if not rows:
        raw.warnings.append("Tidak ada baris data di bawah header.")
    raw.period = _detect_period(raw)
    return raw


def _detect_period(raw: RawData):
    from collections import Counter
    try:
        th, bl = raw.values("Tahun"), raw.values("Bulan")
        hr = raw.values("Hari")
        pairs = Counter((int(_num(a)), int(_num(b)), int(_num(c)))
                        for a, b, c in zip(th, bl, hr) if a and b)
        if pairs:
            y, m, d = pairs.most_common(1)[0][0]
            return dt.date(y, m, min(max(d, 1), calendar.monthrange(y, m)[1]))
    except Exception:
        pass
    try:
        dates = [v for v in raw.values("tgl") if isinstance(v, (dt.date, dt.datetime))]
        if dates:
            d = Counter(dates).most_common(1)[0][0]
            return d.date() if isinstance(d, dt.datetime) else d
    except Exception:
        pass
    return None


def _jenis(rec) -> str:
    return str(rec.get("jenisinvestasi") or "").strip().upper()


def _is_non_mi_credit(j: str) -> bool:
    """Eksposur kredit Non-MI: DOC, Deposito, NCD, Obligasi & Sukuk korporasi.
    Tidak termasuk SBI, SBN (SBP*), saham, reksa dana, penyertaan langsung, tanah/bangunan."""
    if not j or "SBP" in j or j in ("SBI", "SHM", "REKSA", "PL", "TB"):
        return False
    return (j in ("DEP", "DOC", "NCD", "TAB") or j.startswith("OBL") or j.startswith("SUK")
            or "EBA" in j or "MTN" in j)


KNOWN_JENIS_PREFIX = ("DEP", "DOC", "NCD", "TAB", "OBL", "SUK", "SBP", "SBI", "SHM", "REKSA",
                      "PL", "TB", "EBA", "MTN")


def raw_summary(raw: RawData) -> dict:
    by_jenis, unknown = {}, set()
    total = 0.0
    for rec in raw.records():
        j = _jenis(rec)
        nw = _num(rec.get("nilaiwajar"))
        by_jenis[j] = by_jenis.get(j, 0.0) + nw
        total += nw
        if j and not j.startswith(KNOWN_JENIS_PREFIX) and "SBP" not in j:
            unknown.add(j)
    return {"by_jenis": by_jenis, "total_nw": total, "n_rows": len(raw.rows),
            "unknown_jenis": sorted(unknown)}


def credit_non_mi(raw: RawData) -> list:
    """[(pihak, eksposur_miliar, rating_instrumen_unik)] urut eksposur terbesar."""
    agg, rts = {}, {}
    for rec in raw.records():
        if not _is_non_mi_credit(_jenis(rec)):
            continue
        p = str(rec.get("pihak") or rec.get("company") or "(TANPA PIHAK)").strip()
        agg[p] = agg.get(p, 0.0) + _num(rec.get("nilaiwajar"))
        r = rec.get("rating")
        if r:
            rts.setdefault(p, set()).add(str(r).strip())
    out = [(p, miliar_trunc(v), ", ".join(sorted(rts.get(p, [])))) for p, v in agg.items()]
    out.sort(key=lambda x: (-x[1], x[0]))
    return out


def credit_mi(raw: RawData) -> list:
    """[(pihak_MI, nominal_miliar)] dari investasi reksa dana (JenisInvestasi REKSA)."""
    agg, order = {}, []
    for rec in raw.records():
        if _jenis(rec) != "REKSA":
            continue
        p = str(rec.get("pihak") or rec.get("company") or "").strip()
        if p not in agg:
            order.append(p)
        agg[p] = agg.get(p, 0.0) + _num(rec.get("nilaiwajar"))
    return [(p, miliar_trunc(agg[p])) for p in order]


def reksa_dana_list(raw: RawData) -> list:
    out = []
    for rec in raw.records():
        if _jenis(rec) != "REKSA":
            continue
        out.append({"nama": str(rec.get("kode") or "").strip(),
                    "mi": str(rec.get("company") or "").strip(),
                    "jenis": str(rec.get("type") or "").strip(),
                    "np": _num(rec.get("nilaiperolehan")),
                    "nw": _num(rec.get("nilaiwajar"))})
    return out


def spi_rd_per_jenis(raw: RawData) -> dict:
    out = {}
    for rd in reksa_dana_list(raw):
        k = rd["jenis"] or "Lainnya"
        np_, nw = out.get(k, (0.0, 0.0))
        out[k] = (np_ + rd["np"], nw + rd["nw"])
    return out


def spi_saham_existing_juta(raw: RawData) -> float:
    s = 0.0
    for rec in raw.records():
        if _jenis(rec) == "SHM":
            s += _num(rec.get("nilaiwajar")) - _num(rec.get("nilaiperolehan"))
    return s / 1e6


def find_fund_nw(raw: RawData, name: str) -> tuple:
    """Cari Nilai Wajar reksa dana di RAW berdasarkan nama (kolom Kode)."""
    key = _norm(name)
    best = None
    for rd in reksa_dana_list(raw):
        if _norm(rd["nama"]) == key:
            return rd["nama"], rd["nw"]
        if key and (key in _norm(rd["nama"]) or _norm(rd["nama"]) in key):
            best = best or (rd["nama"], rd["nw"])
    return best if best else ("", 0.0)


# --------------------------------------------------------------------------
# Master rating (disimpan supaya bulan depan tidak perlu isi ulang)
# --------------------------------------------------------------------------
def load_master(template_path=TEMPLATE_PATH) -> dict:
    try:
        m = json.loads(MASTER_PATH.read_text(encoding="utf-8"))
        m.setdefault("non_mi", {})
        m.setdefault("mi", [])
        return m
    except Exception:
        return master_from_template(template_path)


def save_master(master: dict):
    MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
    MASTER_PATH.write_text(json.dumps(master, ensure_ascii=False, indent=2), encoding="utf-8")


def master_from_template(template_path=TEMPLATE_PATH) -> dict:
    wb = openpyxl.load_workbook(template_path, data_only=True)
    non_mi = {}
    ws = wb[SH_RATING]
    for r in range(2, ws.max_row + 1):
        n, rt = ws.cell(r, 1).value, ws.cell(r, 3).value
        if n and rt:
            non_mi[str(n).strip().upper()] = str(rt).strip()
    ws = wb[SH_KREDIT]
    for r in range(9, 40):
        n, d, rt = ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value
        if n and isinstance(d, (int, float)) and rt:
            non_mi[str(n).strip().upper()] = str(rt).strip()
    mi = []
    ws = wb[SH_CHK2]
    for r in range(3, 20):
        n, b, rt = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
        if n and isinstance(b, (int, float)):
            mi.append({"pihak": str(n).strip().upper(), "nama": str(n).strip(),
                       "rating": str(rt or "").strip()})
    wb.close()
    return {"non_mi": non_mi, "mi": mi}


def build_credit_rows(raw: RawData, master: dict, pd_kelas: dict) -> list:
    rows = []
    for pihak, eks, _rt in credit_non_mi(raw):
        rating = master.get("non_mi", {}).get(pihak.strip().upper(), "")
        kelas = rating_class(rating) or ""
        rows.append(CreditRow(nama=pihak, eksposur=eks, rating=rating, kelas=kelas,
                              pd=pd_kelas.get(kelas, 0.0) if kelas else 0.0))
    return rows


def build_mi_rows(raw: RawData, master: dict) -> list:
    found = dict(credit_mi(raw))
    keyed = {_norm(k): (k, v) for k, v in found.items()}
    rows, used = [], set()
    for m in master.get("mi", []):
        k = _norm(m.get("pihak"))
        if k in keyed:
            p, v = keyed[k]
            rows.append(MIRow(pihak=p, nama=m.get("nama") or smart_title(p), nominal=v,
                              rating=m.get("rating", "")))
            used.add(k)
    rest = [(p, v) for p, v in found.items() if _norm(p) not in used]
    rest.sort(key=lambda x: -x[1])
    for p, v in rest:
        rows.append(MIRow(pihak=p, nama=smart_title(p), nominal=v, rating=""))
    return rows


def update_master_from_inputs(master: dict, inp: StressInputs) -> dict:
    master = json.loads(json.dumps(master))
    for c in inp.credit:
        if c.nama and c.rating:
            master.setdefault("non_mi", {})[c.nama.strip().upper()] = c.rating.strip()
    mi_list = master.setdefault("mi", [])
    idx = {_norm(m.get("pihak")): m for m in mi_list}
    for m in inp.mi:
        k = _norm(m.pihak or m.nama)
        if k in idx:
            idx[k]["nama"] = m.nama
            idx[k]["rating"] = m.rating
        else:
            mi_list.append({"pihak": (m.pihak or m.nama).upper(), "nama": m.nama, "rating": m.rating})
    return master


# --------------------------------------------------------------------------
# Budget & Realisasi (opsional, untuk isi otomatis)
# --------------------------------------------------------------------------
@dataclass
class BudgetData:
    values: dict = field(default_factory=dict)
    sources: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)


def _find_sheet(wb, name):
    for ws in wb.worksheets:
        if _norm(ws.title) == _norm(name):
            return ws
    raise KeyError(name)


def _date_cols(row):
    cols = {}
    for i, v in enumerate(row):
        if isinstance(v, (dt.datetime, dt.date)):
            cols[(v.year, v.month)] = i
    return cols


def read_budget(src, year: int, month: int) -> BudgetData:
    """Ambil angka dari file Budget & Realisasi (sheet ASUMSI HU & ASUMSI NERACA)."""
    from openpyxl.utils import get_column_letter as gcl
    bd = BudgetData()
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    try:
        ws = _find_sheet(wb, "ASUMSI HU")
        rows = []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_col=40, values_only=True), 1):
            rows.append(list(row))
            if i >= 700:
                break
        # cari section
        sections = {}
        for i, row in enumerate(rows):
            a = row[0] if row else None
            if isinstance(a, str):
                dc = _date_cols(row)
                if len(dc) >= 6:
                    sections.setdefault(a.strip().upper(), (i, dc))

        def sec_val(sec, label, ym):
            if sec not in sections:
                raise KeyError(f"Bagian '{sec}' tidak ditemukan di sheet ASUMSI HU")
            h, dc = sections[sec]
            if ym not in dc:
                raise KeyError(f"Kolom bulan {ym[1]}/{ym[0]} tidak ada di bagian '{sec}'")
            col = dc[ym]
            for j in range(h + 1, min(h + 40, len(rows))):
                r = rows[j]
                if j > h + 1 and len(_date_cols(r)) >= 6:
                    break
                a = r[0] if r else None
                if isinstance(a, str) and _norm(a) == _norm(label):
                    v = r[col] if col < len(r) else None
                    if not isinstance(v, (int, float)):
                        raise ValueError(f"Nilai '{label}' ({sec}) kosong/bukan angka - simpan ulang file "
                                         "Budget di Excel agar nilai rumusnya tersimpan")
                    return float(v), f"ASUMSI HU!{gcl(col + 1)}{j + 1}"
            raise KeyError(f"Baris '{label}' tidak ditemukan di bagian '{sec}'")

        rep, dec = (year, month), (year, 12)
        spec = {
            "ex_total_aset": ("RATA-RATA NILAI WAJAR", "Total", rep, 1 / 1000),
            "ex_spi": ("SPI", "Total", rep, 1 / 1000),
            "ex_laba": ("PENDAPATAN INVESTASI BERSIH YTD", "Total", rep, 1 / 1000),
            "ex_delta_spi": ("DELTA SPI YTD", "Total", rep, 1 / 1000),
            "pr_total_aset": ("RATA-RATA NILAI WAJAR", "Total", dec, 1 / 1000),
            "pr_spi": ("SPI", "Total", dec, 1 / 1000),
            "pr_laba": ("PENDAPATAN INVESTASI BERSIH YTD", "Total", dec, 1 / 1000),
            "pr_delta_spi": ("DELTA SPI YTD", "Total", dec, 1 / 1000),
            "spi_saham_des_juta": ("SPI", "Saham", dec, 1.0),
            "spi_saham_rep_juta": ("SPI", "Saham", rep, 1.0),
        }
        for key, (sec, lab, ym, mult) in spec.items():
            try:
                v, srcref = sec_val(sec, lab, ym)
                bd.values[key] = v * mult
                bd.sources[key] = srcref
            except Exception as e:
                bd.warnings.append(f"{key}: {e}")
        # Delta SPI selain saham (untuk kalkulator yield): SPI Total - Saham - PL - Tanah Bangunan
        try:
            parts = {}
            for ym, tag in ((rep, "rep"), (dec, "dec")):
                tot, _ = sec_val("SPI", "Total", ym)
                sah, _ = sec_val("SPI", "Saham", ym)
                pl, _ = sec_val("SPI", "Penyertaan Langsung", ym)
                tb, _ = sec_val("SPI", "Tanah Bangunan", ym)
                parts[tag] = tot - sah - pl - tb
            bd.values["yc_delta_spi_raw"] = (parts["dec"] - parts["rep"]) / 1000
            bd.values["yc_delta_spi"] = round_half_up(bd.values["yc_delta_spi_raw"], 3)
            bd.sources["yc_delta_spi"] = ("ASUMSI HU bagian SPI: (Total - Saham - Penyertaan Langsung - "
                                          "Tanah Bangunan) Des dikurangi bulan laporan")
        except Exception as e:
            bd.warnings.append(f"yc_delta_spi: {e}")
    except KeyError as e:
        bd.warnings.append(str(e))
    # Aset Neto
    try:
        ws = _find_sheet(wb, "ASUMSI NERACA")
        hdr_cols, found = None, None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_col=40, values_only=True), 1):
            row = list(row)
            if hdr_cols is None and len(_date_cols(row)) >= 6:
                hdr_cols = _date_cols(row)
            a = row[0] if row else None
            if hdr_cols and isinstance(a, str) and _norm(a) == _norm("Total Aset Neto"):
                col = hdr_cols.get((year, month))
                if col is not None and isinstance(row[col], (int, float)):
                    found = (row[col], f"ASUMSI NERACA!{openpyxl.utils.get_column_letter(col + 1)}{i}")
                break
            if i > 400:
                break
        if found:
            bd.values["aset_neto"] = round_half_up(found[0] * 1e6)
            bd.sources["aset_neto"] = found[1]
        else:
            bd.warnings.append("Total Aset Neto bulan laporan tidak ditemukan di ASUMSI NERACA")
    except KeyError:
        bd.warnings.append("Sheet ASUMSI NERACA tidak ditemukan")
    wb.close()
    return bd


# --------------------------------------------------------------------------
# PDF "Laporan SPI Reksadana per Jenis" (DISK)
# --------------------------------------------------------------------------
SPI_RD_JENIS = ["Campuran", "Indeks", "Pasar Uang", "Pendapatan Tetap", "Saham", "Terproteksi", "Total"]
_NUM_ID = r"\(?-?\d{1,3}(?:\.\d{3})*,\d+\)?"


def _id_number(tok: str) -> float:
    """'(3.040.594.067,36)' -> -3040594067.36 ; '500.762.376,57' -> 500762376.57"""
    neg = tok.startswith("(") or tok.startswith("-")
    v = float(tok.strip("()-").replace(".", "").replace(",", "."))
    return -v if neg else v


def read_spi_reksadana_pdf(data: bytes) -> dict:
    """Baca tabel Laporan SPI Reksadana per Jenis.
    Hasil: {"rows": {jenis: {"unit","np","nw","spi"}}, "periode": date|None, "warnings": [...]}"""
    from ffs_reader import pdf_text, _MONTH_WORDS
    text = pdf_text(data)
    flat = re.sub(r"\s+", " ", text)
    rows, warns = {}, []
    for jenis in SPI_RD_JENIS:
        pat = r"(?<![A-Za-z])" + r"\s+".join(map(re.escape, jenis.split())) + r"\s+((?:" + _NUM_ID + r"\s*){1,4})"
        m = re.search(pat, flat)
        if not m:
            continue
        nums = [_id_number(x) for x in re.findall(_NUM_ID, m.group(1))]
        nums += [0.0] * (4 - len(nums))
        unit, np_, nw, spi = nums[:4]
        rows[jenis] = {"unit": unit, "np": np_, "nw": nw, "spi": spi}
        if nw and abs((nw - np_) - spi) > 1:
            warns.append(f"Baris {jenis}: SPI ({spi:,.2f}) tidak sama dengan Nilai Wajar - Nilai Perolehan.")
    periode = None
    m = re.search(r"Periode\s*:?\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", flat)
    if m and _MONTH_WORDS.get(m.group(2).lower()):
        try:
            periode = dt.date(int(m.group(3)), _MONTH_WORDS[m.group(2).lower()], int(m.group(1)))
        except ValueError:
            periode = None
    if "Saham" not in rows:
        warns.append("Baris 'Saham' tidak ditemukan di PDF. Pastikan file adalah Laporan SPI Reksadana per Jenis.")
    return {"rows": rows, "periode": periode, "warnings": warns}


# --------------------------------------------------------------------------
# Nilai awal dari template (file bulan sebelumnya)
# --------------------------------------------------------------------------
def _parse_pct_formula(f) -> tuple:
    """'=47.91%+5.55%' -> (47.91, 5.55)."""
    if isinstance(f, (int, float)):
        return float(f) * 100, 0.0
    nums = [float(x) for x in re.findall(r"(-?\d+(?:\.\d+)?)\s*%", str(f))]
    if not nums:
        return 0.0, 0.0
    if len(nums) == 1:
        return nums[0], 0.0
    return nums[0], sum(nums[1:])


def defaults_from_template(template_path=TEMPLATE_PATH) -> StressInputs:
    wv = openpyxl.load_workbook(template_path, data_only=True)
    wf = openpyxl.load_workbook(template_path)
    h = wv[SH_HASIL]
    inp = StressInputs()
    d = wv[SH_PASAR_ESC]["B2"].value
    if isinstance(d, (dt.datetime, dt.date)):
        inp.tahun, inp.bulan = d.year, d.month
    g = lambda ws, a: _num(ws[a].value)
    inp.ihsg_existing, inp.ihsg_escalated, inp.ihsg_severe = g(h, "C4"), g(h, "E4"), g(h, "F4")
    inp.yield_existing, inp.yield_escalated, inp.yield_severe = g(h, "C5"), g(h, "E5"), g(h, "F5")
    inp.kurs_existing, inp.kurs_baseline = g(h, "C6"), g(h, "D6")
    inp.kurs_escalated, inp.kurs_severe = g(h, "E6"), g(h, "F6")
    inp.ex_total_aset, inp.ex_spi, inp.ex_laba, inp.ex_delta_spi = (g(h, "C10"), g(h, "C11"),
                                                                    g(h, "C12"), g(h, "C13"))
    inp.pr_total_aset, inp.pr_spi, inp.pr_laba, inp.pr_delta_spi = (g(h, "L4"), g(h, "L5"),
                                                                    g(h, "L6"), g(h, "L7"))
    inp.aset_neto = g(h, "L12")
    sb, sbf = wv[SH_SB], wf[SH_SB]
    inp.yc_nilai_wajar = g(sb, "B5")
    inp.yc_delta_spi = g(sb, "F5")
    inp.yc_tenor = int(g(sb, "B8") or 10)
    k = wv[SH_KRD]
    inp.rd_saham_np, inp.rd_saham_nw = g(k, "B14"), g(k, "C14")
    inp.rd_campuran_np, inp.rd_campuran_nw = g(k, "B15"), g(k, "C15")
    l8_nums = [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?", str(sbf["L8"].value or ""))]
    rd_juta = (inp.rd_saham_nw - inp.rd_saham_np) / 1e6
    inp.spi_saham_des_juta = l8_nums[0] if l8_nums else 0.0
    inp.rd_spi_saham_juta = l8_nums[1] if len(l8_nums) > 1 else round_half_up(rd_juta, 2)
    inp.spi_saham_rep_juta = g(sb, "L9") - inp.rd_spi_saham_juta
    kf = wf[SH_KRD]
    funds = []
    r = 4
    while r < 12 and k.cell(r, 1).value:
        obl, pu = _parse_pct_formula(kf.cell(r, 3).value)
        funds.append(FundSlot(nama=str(k.cell(r, 1).value).strip(),
                              kode_raw=str(k.cell(r, 1).value).strip(),
                              saham=round_half_up(_num(k.cell(r, 2).value) * 100, 4),
                              obligasi=obl, pasar_uang=pu,
                              nilai_wajar=_num(k.cell(r, 7).value)))
        r += 1
    inp.funds = funds
    lik = wv[SH_LIK_ESC]
    inp.likuiditas_eksposur = g(lik, "D7")
    inp.likuiditas_tenor = int(g(lik, "D9") or 10)
    kr = wv[SH_KREDIT]
    inp.lgd = _num(kr["G9"].value, DEFAULT_LGD) or DEFAULT_LGD
    credit = []
    for r in range(9, 40):
        n, dd, rt, pdv = kr.cell(r, 3).value, kr.cell(r, 4).value, kr.cell(r, 5).value, kr.cell(r, 6).value
        if n and isinstance(dd, (int, float)):
            kel = rating_class(rt) or ""
            credit.append(CreditRow(nama=str(n), eksposur=float(dd), rating=str(rt or ""),
                                    kelas=kel, pd=_num(pdv)))
    inp.credit = credit
    mi = []
    c2 = wv[SH_CHK2]
    for r in range(3, 20):
        n, b, rt = c2.cell(r, 1).value, c2.cell(r, 2).value, c2.cell(r, 3).value
        if n and isinstance(b, (int, float)):
            mi.append(MIRow(pihak=str(n).upper(), nama=str(n), nominal=float(b), rating=str(rt or "")))
    inp.mi = mi
    wv.close()
    wf.close()
    return inp


# --------------------------------------------------------------------------
# Validasi sebelum membuat Excel
# --------------------------------------------------------------------------
def validate_inputs(inp: StressInputs, raw: RawData | None) -> tuple:
    errors, warns = [], []
    if raw is None or not raw.rows:
        errors.append("Data RAW Investasi (file Rekap Investasi) belum dimasukkan.")
    elif raw.period and (raw.period.year, raw.period.month) != (inp.tahun, inp.bulan):
        warns.append(f"Periode di Rekap Investasi ({BULAN[raw.period.month]} {raw.period.year}) "
                     f"berbeda dengan periode yang dipilih ({BULAN[inp.bulan]} {inp.tahun}).")
    b = compute_baseline(inp)
    if not inp.ihsg_existing:
        errors.append("IHSG existing belum diisi.")
    if not b["ihsg_escalated_valid"]:
        lo, hi = b["ihsg_range"]
        errors.append(f"IHSG Escalated ({inp.ihsg_escalated:,.2f}) harus berada di antara IHSG Severe "
                      f"dan IHSG Baseline ({lo:,.2f} – {hi:,.2f}).")
    if not inp.yield_existing:
        errors.append("Yield SBN 10 tahun existing belum diisi.")
    if not inp.yc_nilai_wajar:
        errors.append("Nilai Wajar kalkulator yield belum diisi.")
    if math.isnan(b["yield_baseline"]):
        errors.append("Yield baseline tidak bisa dihitung (cek Nilai Wajar & Delta SPI).")
    if not inp.aset_neto:
        errors.append("Aset Neto belum diisi.")
    for nm, v in (("Total Aset Investasi proyeksi", inp.pr_total_aset),
                  ("Total Aset Investasi existing", inp.ex_total_aset)):
        if not v:
            errors.append(f"{nm} belum diisi.")
    if not inp.funds:
        errors.append("Komposisi RD: belum ada reksa dana campuran.")
    for f in inp.funds:
        tot = f.saham + f.obligasi + f.pasar_uang
        if not f.nama:
            errors.append("Komposisi RD: ada baris reksa dana tanpa nama.")
        elif tot == 0:
            errors.append(f"Komposisi '{f.nama}' belum diisi (ambil dari Fund Fact Sheet).")
        elif abs(tot - 100) > 1.0:
            warns.append(f"Komposisi '{f.nama}' berjumlah {tot:.2f}% (bukan ±100%).")
        if f.nama and not f.nilai_wajar:
            warns.append(f"Nilai Wajar '{f.nama}' = 0 (tidak ditemukan di RAW?).")
    if len(inp.funds) > 6:
        errors.append("Komposisi RD maksimal 6 reksa dana.")
    if not inp.rd_saham_nw and not inp.rd_saham_np:
        warns.append("SPI Reksa Dana Saham (NP/NW) di tab 4 masih 0.")
    if not inp.rd_spi_saham_juta:
        warns.append("SPI Reksa Dana Saham (juta) di Calculator IHSG Skenario Baseline masih 0.")
    else:
        rd_tab4 = (inp.rd_saham_nw - inp.rd_saham_np) / 1e6
        if (inp.rd_saham_nw or inp.rd_saham_np) and abs(inp.rd_spi_saham_juta - rd_tab4) > 1:
            warns.append(f"SPI Reksa Dana Saham di kalkulator IHSG ({inp.rd_spi_saham_juta:,.2f} juta) berbeda "
                         f"dengan SPI Saham di tab 4 ({rd_tab4:,.2f} juta). Cek lagi PDF Laporan SPI Reksadana.")
    for c in inp.credit:
        if not c.rating:
            errors.append(f"Rating grup '{c.nama}' belum diisi (tab Risiko Kredit).")
        elif c.kelas not in DEFAULT_PD_KELAS:
            errors.append(f"Kelas rating '{c.nama}' ({c.rating}) belum dipilih.")
    if len(inp.credit) > 30:
        errors.append("Risiko Kredit Non-MI maksimal 30 grup (batas tabel template).")
    for m in inp.mi:
        if m.rating not in inp.pd_mi:
            errors.append(f"Rating MI '{m.nama}' belum dipilih (Sangat Bagus/Bagus/Cukup).")
    if len(inp.mi) > 13:
        errors.append("Risiko Kredit MI maksimal 13 manajer investasi.")
    for key, nm in (("esc", "Escalated"), ("sev", "Severe")):
        if getattr(inp, f"b_manual_{key}"):
            bv = float(getattr(inp, f"b_{key}"))
            if bv == 0:
                errors.append(f"Peningkatan suku bunga {nm} (manual) belum diisi (tab 2).")
            elif bv < 0:
                warns.append(f"Peningkatan suku bunga {nm} bernilai negatif ({bv * 100:.4f}%).")
        elif not getattr(inp, f"yield_{'escalated' if key == 'esc' else 'severe'}"):
            errors.append(f"Yield {nm} belum diisi (tab 2).")
    return errors, warns


# --------------------------------------------------------------------------
# Pembuatan workbook
# --------------------------------------------------------------------------
@dataclass
class BuildResult:
    content: bytes
    filename: str
    values: dict
    summary: dict
    warnings: list


def _set(ws, addr, value):
    ws[addr].value = value


def _replace_paren(text: str, new: str) -> str:
    if not isinstance(text, str):
        return text
    if "(" in text:
        return re.sub(r"\([^)]*\)", f"({new})", text, count=1)
    return f"{text} ({new})"


def _replace_year(text: str, year: int) -> str:
    if not isinstance(text, str):
        return text
    return re.sub(r"(19|20)\d{2}", str(year), text)


def _set_comment(ws, addr, text):
    old = ws[addr].comment
    author = old.author if old is not None and old.author else "Manajemen Risiko"
    c = Comment(text, author)
    c.width, c.height = 160, 48
    ws[addr].comment = c


def _clear_pivots(ws):
    try:
        ws._pivots = []
    except Exception:
        pass


def build_workbook(inp: StressInputs, raw: RawData, template_path=TEMPLATE_PATH) -> BuildResult:
    warns = []
    wb = openpyxl.load_workbook(template_path)
    rep = month_end(inp.tahun, inp.bulan)
    rep_dt = dt.datetime(rep.year, rep.month, rep.day)
    base = compute_baseline(inp)
    y_base = base["yield_baseline"]
    shift_esc, shift_sev = base["shift_escalated"], base["shift_severe"]

    # ---------------- RAW INVESTASI ----------------
    ws = wb[SH_RAW]
    ncol = len(raw.headers)
    fmts = [ws.cell(2, c).number_format for c in range(1, ncol + 1)]
    styles = [copy(ws.cell(2, c)._style) for c in range(1, ncol + 1)]
    old_max = ws.max_row
    for r in range(2, old_max + 1):
        for c in range(1, ncol + 1):
            ws.cell(r, c).value = None
    for i, row in enumerate(raw.rows):
        r = i + 2
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            v = row[c - 1]
            # Seperti "paste values" manual: format sel yang sudah ada dipertahankan,
            # baris baru (melebihi baris lama) memakai format baris 2.
            fmt_before = cell.number_format
            cell.value = v
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"   # teks yang diawali '=' jangan jadi rumus
            if r > old_max:
                cell._style = copy(styles[c - 1])
                cell.number_format = fmts[c - 1]
            else:
                cell.number_format = fmt_before
    # rapikan sisa baris lama
    if old_max > len(raw.rows) + 1:
        ws.delete_rows(len(raw.rows) + 2, old_max - len(raw.rows) - 1)
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(ncol)}{len(raw.rows) + 1}"

    # ---------------- Hasil Stress Testing ----------------
    ws = wb[SH_HASIL]
    tgl_singkat = f"{rep.day} {BULAN_SINGKAT[rep.month]} {rep.year}"
    ws["B1"].value = f"*) Baseline = Proyeksi 31 Desember {rep.year}"
    for a in ("C3", "C21", "C39"):
        ws[a].value = _replace_paren(ws[a].value, tgl_singkat)
    ws["K3"].value = _replace_year(ws["K3"].value, rep.year)
    for a, v in (("C4", inp.ihsg_existing), ("E4", inp.ihsg_escalated), ("F4", inp.ihsg_severe),
                 ("C5", inp.yield_existing), ("E5", base["yield_escalated"]), ("F5", base["yield_severe"]),
                 ("C6", inp.kurs_existing), ("D6", inp.kurs_baseline), ("E6", inp.kurs_escalated),
                 ("F6", inp.kurs_severe),
                 ("C10", inp.ex_total_aset), ("C11", inp.ex_spi), ("C12", inp.ex_laba),
                 ("C13", inp.ex_delta_spi),
                 ("L4", inp.pr_total_aset), ("L5", inp.pr_spi), ("L6", inp.pr_laba),
                 ("L7", inp.pr_delta_spi), ("L12", inp.aset_neto)):
        ws[a].value = v
    _set_comment(ws, "C3", f"Based on data Aktual {MONTH_EN[rep.month]} {rep.year}")
    _set_comment(ws, "D3", f"Based on data Proyeksi Dec {rep.year}")
    _set_comment(ws, "E3", "Skenario Adjustment dari Tim MR")
    _set_comment(ws, "F3", "Skenario berdasarkan ketentuan terbaru OJK")
    _set_comment(ws, "C21", f"Based on data Aktual {MONTH_EN[rep.month]} {rep.year}")
    _set_comment(ws, "C39", f"Based on data Aktual {MONTH_EN[rep.month]} {rep.year}")

    # ---------------- Skenario Baseline ----------------
    ws = wb[SH_SB]
    tenor = int(inp.yc_tenor or 10)
    ws["B5"].value = inp.yc_nilai_wajar
    ws["B9"].value = inp.yc_nilai_wajar
    ws["B8"].value = tenor
    ws["B10"].value = inp.yield_existing
    ws["C5"].value = base["pv_skenario"]
    ws["D5"].value = base["pv_existing"]
    ws["F4"].value = inp.yc_delta_spi
    ws["F5"].value = inp.yc_delta_spi
    ws["L7"].value = inp.ihsg_existing
    ws["L8"].value = f"=({int(base['L8a'])})+({int(base['L8b'])})"
    ws["L9"].value = int(base["L9"])

    # ---------------- Komposisi RD ----------------
    _write_komposisi(wb, inp, warns)

    # ---------------- Risiko Pasar (4 sheet) ----------------
    for name, shift in ((SH_PASAR_ESC, shift_esc), (SH_PASAR_SEV, shift_sev),
                        (SH_PASAR_BASE, 0.0), (SH_PASAR_EXIST, 0.0)):
        ws = wb[name]
        ws["B2"].value = rep_dt
        for col in "EFGHIJKLM":
            if name in (SH_PASAR_ESC, SH_PASAR_SEV):
                ws[f"{col}18"].value = shift
            if name == SH_PASAR_EXIST:
                ws[f"{col}17"].value = inp.yield_existing
        if name in (SH_PASAR_ESC, SH_PASAR_SEV):
            ws["B24"].value = f"Harga Saham per 31 Desember {rep.year}"
        else:
            ws["B24"].value = f"Harga Saham per {rep.day} {BULAN[rep.month]} {rep.year}"

    # ---------------- Risiko Kredit ----------------
    _write_kredit(wb, inp, rep_dt, warns)
    _write_check1(wb, inp)
    _write_check2(wb, inp)
    _write_rating_sheet(wb, inp)

    # ---------------- Risiko Likuiditas ----------------
    lp = liquidity_params(inp, base)
    for name, key in ((SH_LIK_ESC, "esc"), (SH_LIK_SEV, "sev")):
        ws = wb[name]
        ws["D7"].value = inp.likuiditas_eksposur
        ws["D9"].value = int(inp.likuiditas_tenor or 10)
        ws["D10"].value = lp[key]["a"]      # (a) Peningkatan suku bunga = B skenario
        # (b) Yield awal: D11 tetap rumus ='Hasil Stress Testing'!D41 seperti template

    # ---------------- CHECK (tersembunyi) ----------------
    _write_check_hidden(wb, raw)

    # Rumus bucket tenor >10 th (L9, M9, L11, M11) dan Hasil!C40 versi terbaru
    _apply_formula_fixes(wb)

    # ---------------- Properti workbook ----------------
    wb.properties.title = f"Stress Test Risiko {BULAN[rep.month]} {rep.year}"
    wb.properties.modified = dt.datetime.now()
    wb.calculation.fullCalcOnLoad = True

    # ---------------- Hitung semua rumus ----------------
    eng = Engine(wb)
    values = eng.evaluate_all()
    errs = [(s, openpyxl.utils.get_column_letter(c) + str(r), v) for (s, r, c), v in values.items()
            if isinstance(v, XLErr)]
    for s, a, v in errs[:15]:
        warns.append(f"Rumus menghasilkan {v} di {s}!{a}")

    buf = io.BytesIO()
    wb.save(buf)
    content = inject_cached_values(buf.getvalue(), values)
    summary = make_summary(values, base, eng)
    return BuildResult(content=content, filename=output_filename(inp.tahun, inp.bulan),
                       values=values, summary=summary, warnings=warns)


FIXED_FORMULAS = {
    "L9": ("=SUM(SUMIFS('RAW INVESTASI'!W:W,'RAW INVESTASI'!I:I,\">\" & EDATE($B$2,120),"
           "'RAW INVESTASI'!I:I,\"<=\" & EDATE($B$2,180),'RAW INVESTASI'!N:N,"
           "{\"DEP\",\"DOC\",\"NCD\",\"SBI\"}))/1000000000"),
    "M9": ("=SUM(SUMIFS('RAW INVESTASI'!W:W,'RAW INVESTASI'!I:I,\">\" & EDATE($B$2,180),"
           "'RAW INVESTASI'!N:N,{\"DEP\",\"DOC\",\"NCD\",\"SBI\"}))/1000000000"),
    "L11": ("=SUM(SUMIFS('RAW INVESTASI'!W:W,'RAW INVESTASI'!I:I,\">\" & EDATE($B$2,120),"
            "'RAW INVESTASI'!I:I,\"<=\" & EDATE($B$2,180),'RAW INVESTASI'!N:N,{\"*OBL*\",\"*SUK*\"},"
            "'RAW INVESTASI'!N:N,\"<>*SBP*\"))/1000000000"),
    "M11": ("=SUM(SUMIFS('RAW INVESTASI'!W:W,'RAW INVESTASI'!I:I,\">\" & EDATE($B$2,180),"
            "'RAW INVESTASI'!N:N,{\"*OBL*\",\"*SUK*\"},'RAW INVESTASI'!N:N,\"<>*SBP*\"))/1000000000"),
}


def _apply_formula_fixes(wb):
    from openpyxl.worksheet.formula import ArrayFormula
    for name in (SH_PASAR_ESC, SH_PASAR_SEV, SH_PASAR_BASE, SH_PASAR_EXIST):
        ws = wb[name]
        for addr, f in FIXED_FORMULAS.items():
            ws[addr].value = ArrayFormula(addr, f)
    wb[SH_HASIL]["C40"].value = "=C4"


def _write_komposisi(wb, inp: StressInputs, warns):
    ws = wb[SH_KRD]
    funds = inp.funds
    n = len(funds)
    if n < 1:
        return
    if n > 6:
        raise ValueError("Komposisi RD maksimal 6 reksa dana")
    st = {r: [copy(ws.cell(r, c)._style) for c in range(1, 8)] for r in (4, 5, 6, 7)}
    st_empty = [copy(ws.cell(10, c)._style) for c in range(1, 8)]
    h4, h5 = ws.row_dimensions[4].height, ws.row_dimensions[5].height
    sum_row = 4 + n
    if n != 2:
        for r in range(4, 12):
            for c in range(1, 8):
                ws.cell(r, c).value = None
                ws.cell(r, c)._style = copy(st_empty[c - 1])
            ws.row_dimensions[r].height = None
    for i, f in enumerate(funds):
        r = 4 + i
        if n != 2:
            src = st[4] if i == 0 else st[5]
            for c in range(1, 8):
                ws.cell(r, c)._style = copy(src[c - 1])
            ws.row_dimensions[r].height = h4 if i == 0 else h5
        ws.cell(r, 1).value = f.nama
        ws.cell(r, 2).value = round_half_up(f.saham / 100, 8)
        ws.cell(r, 3).value = f"={fmt_pct_literal(f.obligasi)}%+{fmt_pct_literal(f.pasar_uang)}%"
        ws.cell(r, 4).value = f"=B{r}+C{r}"
        if i == 0:
            ws.cell(r, 5).value = f"=B{r}*G${r}"
            ws.cell(r, 6).value = f"=C{r}*$G${r}"
        else:
            ws.cell(r, 5).value = f"=B{r}*G${r}"
            ws.cell(r, 6).value = f"=C{r}*G${r}"
        ws.cell(r, 7).value = f.nilai_wajar
    last = 3 + n
    if n != 2:
        for c in range(1, 8):
            ws.cell(sum_row, c)._style = copy(st[6][c - 1])
            ws.cell(sum_row + 1, c)._style = copy(st[7][c - 1])
        for col in "BCDEFG":
            ws[f"{col}{sum_row}"].value = f"=SUM({col}4:{col}{last})"
        ws[f"B{sum_row + 1}"].value = f"=B{sum_row}/D{sum_row}"
        ws[f"C{sum_row + 1}"].value = f"=C{sum_row}/D{sum_row}"
        # perbarui referensi 'Komposisi RD'!E6 di sheet Risiko Pasar
        for name in (SH_PASAR_ESC, SH_PASAR_SEV, SH_PASAR_BASE, SH_PASAR_EXIST):
            w = wb[name]
            for a in ("E13", "E14"):
                v = w[a].value
                if isinstance(v, str):
                    w[a].value = v.replace("'Komposisi RD'!E6", f"'Komposisi RD'!E{sum_row}")
    ws["B14"].value = inp.rd_saham_np
    ws["C14"].value = inp.rd_saham_nw
    ws["B15"].value = inp.rd_campuran_np
    ws["C15"].value = inp.rd_campuran_nw


def _write_kredit(wb, inp: StressInputs, rep_dt, warns):
    ws = wb[SH_KREDIT]
    ws["B2"].value = rep_dt
    rows = inp.credit
    n = len(rows)
    if n > 30:
        raise ValueError("Maksimal 30 grup Non-MI")
    cols = range(2, 10)
    st_norm = {c: copy(ws.cell(9, c)._style) for c in cols}
    st_def = {c: copy(ws.cell(26, c)._style) for c in cols}      # baris rating D (kuning)
    st_blank = {c: copy(ws.cell(29, c)._style) for c in cols}
    st_note = {c: copy(ws.cell(31, c)._style) for c in cols}
    note_text = ws["C31"].value
    note_h = ws.row_dimensions[31].height
    last = 8 + n
    # jenis gaya tiap baris template: 'D' (rating D, sel kuning) atau 'N' (normal)
    def row_kind(r):
        fl = ws.cell(r, 3).fill
        if not fl or fl.fill_type is None:
            return None
        return "D" if ws.cell(r, 6).fill.fgColor.rgb == "FFFFFF00" else "N"
    kinds = {r: row_kind(r) for r in range(9, 40)}
    # baris data
    for i, cr in enumerate(rows):
        r = 9 + i
        want = "D" if cr.kelas == "D" else "N"
        if kinds.get(r) != want:          # gaya template dipertahankan bila jenisnya sama
            src = st_def if want == "D" else st_norm
            for c in cols:
                ws.cell(r, c)._style = copy(src[c])
        ws.cell(r, 2).value = i + 1
        ws.cell(r, 3).value = cr.nama
        ws.cell(r, 4).value = cr.eksposur
        ws.cell(r, 5).value = cr.rating
        ws.cell(r, 6).value = cr.pd
        ws.cell(r, 7).value = inp.lgd
        ws.cell(r, 8).value = f"=D{r}"
        ws.cell(r, 9).value = f"=F{r}*G{r}*H{r}"
        if r == 31 and note_h:
            ws.row_dimensions[r].height = 15.75
    # baris template lama yang kini kosong (sebelum baris 29)
    for r in range(last + 1, 29):
        for c in cols:
            ws.cell(r, c).value = None
            ws.cell(r, c)._style = copy(st_blank[c])
    # catatan Waskita (baris 31) dipindah bila tertimpa data
    if last >= 31:
        note_row = last + 1
        if note_row <= 39:
            for c in cols:
                ws.cell(note_row, c)._style = copy(st_note[c])
                ws.cell(note_row, c).value = None
            ws.cell(note_row, 3).value = note_text
            ws.row_dimensions[note_row].height = note_h
    if last >= 30:
        for r in range(last + 1, 40):
            if ws.cell(r, 8).value and str(ws.cell(r, 8).value).startswith("=D"):
                ws.cell(r, 8).value = None
                ws.cell(r, 9).value = None
    ws.auto_filter.ref = f"B8:I{max(last + 1, 9)}"


def _write_check1(wb, inp: StressInputs):
    """Ganti pivot '2-CHECK (1)' dengan tabel statis berumus (tampilan sama)."""
    ws = wb[SH_CHK1]
    _clear_pivots(ws)
    for r in range(1, 45):
        for c in (1, 2):
            ws.cell(r, c).value = None
            ws.cell(r, c)._style = StyleArray()
    labels, cls = [], {}
    for cr in inp.credit:
        lab = (cr.rating or "").strip()
        if lab and lab not in labels:
            labels.append(lab)
            cls[lab] = cr.kelas
    labels.sort(key=rating_sort_key)
    blue = "FF9BC2E6"
    hdr_fill = PatternFill("solid", fgColor="FFDDEBF7")
    thin = Side(style="thin", color=blue)
    ws["A1"].value = "Row Labels"
    ws["B1"].value = "Nominal"
    for a in ("A1", "B1"):
        ws[a].font = Font(name="Calibri", size=11, bold=True)
        ws[a].fill = hdr_fill
        ws[a].border = Border(bottom=thin)
    ws["A1"].alignment = Alignment(horizontal="left")
    ws["B1"].alignment = Alignment(horizontal="center")
    rows_by_cls = {"AAA": [], "AA": [], "A": [], "D": []}
    for i, lab in enumerate(labels):
        r = 2 + i
        ws.cell(r, 1).value = lab
        ws.cell(r, 1).alignment = Alignment(horizontal="left")
        ws.cell(r, 2).value = (f"=SUMIF('{SH_KREDIT}'!$E$9:$E$39,A{r},'{SH_KREDIT}'!$D$9:$D$39)")
        ws.cell(r, 2).number_format = "0.00"
        if cls.get(lab) in rows_by_cls:
            rows_by_cls[cls[lab]].append(r)
    k = len(labels)
    br, tr = 2 + k, 3 + k
    ws.cell(br, 1).value = "(blank)"
    ws.cell(br, 1).alignment = Alignment(horizontal="left")
    ws.cell(tr, 1).value = "Grand Total"
    ws.cell(tr, 2).value = f"=SUM(B2:B{br})" if k else "=0"
    ws.cell(tr, 2).number_format = "0.00"
    for c in (1, 2):
        cell = ws.cell(tr, c)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.fill = hdr_fill
        cell.border = Border(top=thin)
    ws.cell(tr, 1).alignment = Alignment(horizontal="left")
    for row_e, kel in ((3, "AAA"), (4, "AA"), (5, "A"), (6, "D")):
        refs = rows_by_cls[kel]
        ws.cell(row_e, 5).value = ("=" + "+".join(f"B{r}" for r in refs)) if refs else "=0"
        ws.cell(row_e, 6).value = inp.pd_kelas.get(kel, DEFAULT_PD_KELAS[kel])
        ws.cell(row_e, 7).value = inp.lgd


def _write_check2(wb, inp: StressInputs):
    ws = wb[SH_CHK2]
    st_row = [copy(ws.cell(3, c)._style) for c in (1, 2, 3)]
    old_last = 2
    for r in range(3, 17):
        if ws.cell(r, 1).value not in (None, ""):
            old_last = r
    m = len(inp.mi)
    for r in range(3, max(old_last, 2 + m) + 1):
        for c in (1, 2, 3):
            ws.cell(r, c).value = None
    groups = {"Sangat Bagus": [], "Bagus": [], "Cukup": []}
    for i, mi in enumerate(inp.mi):
        r = 3 + i
        for j, c in enumerate((1, 2, 3)):
            if r > 7:
                ws.cell(r, c)._style = copy(st_row[j])
        ws.cell(r, 1).value = mi.nama
        ws.cell(r, 2).value = mi.nominal
        ws.cell(r, 3).value = mi.rating
        if mi.rating in groups:
            groups[mi.rating].append(r)
    for row_f, key in ((3, "Sangat Bagus"), (4, "Bagus"), (5, "Cukup")):
        refs = groups[key]
        if not refs:
            ws.cell(row_f, 6).value = None
        elif len(refs) == 1:
            ws.cell(row_f, 6).value = f"=B{refs[0]}"
        elif len(refs) == 2:
            ws.cell(row_f, 6).value = f"=B{refs[0]}+B{refs[1]}"
        else:
            ws.cell(row_f, 6).value = "=SUM(" + ",".join(f"B{r}" for r in refs) + ")"
        ws.cell(row_f, 7).value = inp.pd_mi.get(key, DEFAULT_PD_MI[key])
        ws.cell(row_f, 8).value = inp.lgd


def _write_rating_sheet(wb, inp: StressInputs):
    ws = wb[SH_RATING]
    for r in range(2, max(ws.max_row, 2 + len(inp.credit)) + 1):
        for c in (1, 2, 3):
            ws.cell(r, c).value = None
    for i, cr in enumerate(inp.credit):
        ws.cell(2 + i, 1).value = cr.nama
        ws.cell(2 + i, 2).value = cr.eksposur
        ws.cell(2 + i, 3).value = cr.rating


CHECK_ORDER = ["DEP", "DOC", "OBLAFS", "OBLHTM", "PL", "REKSA", "SBPAFS", "SBPHTM", "SHM",
               "SUKSBPAFS", "SUKSBPHTM", "TB", "SUKHTM", "SBI"]


def _write_check_hidden(wb, raw: RawData):
    ws = wb[SH_CHECK]
    _clear_pivots(ws)
    for r in range(1, max(ws.max_row, 40) + 1):
        for c in (1, 2):
            ws.cell(r, c).value = None
    present = []
    for rec in raw.records():
        j = str(rec.get("jenisinvestasi") or "").strip()
        if j and j not in present:
            present.append(j)
    order = [j for j in CHECK_ORDER if j in present] + sorted(j for j in present if j not in CHECK_ORDER)
    ws["A1"].value = "Row Labels"
    ws["B1"].value = "Sum of NilaiWajar"
    for i, j in enumerate(order):
        r = 2 + i
        ws.cell(r, 1).value = j
        ws.cell(r, 2).value = f"=SUMIF('{SH_RAW}'!$N:$N,A{r},'{SH_RAW}'!$W:$W)"
    tr = 2 + len(order)
    ws.cell(tr, 1).value = "Grand Total"
    ws.cell(tr, 2).value = f"=SUM(B2:B{tr - 1})"


# --------------------------------------------------------------------------
# Menyisipkan nilai hasil hitung ke XML (cached value)
# --------------------------------------------------------------------------
_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
       "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
_CELL_F_RE = re.compile(
    r'<c r="(?P<col>[A-Z]+)(?P<row>\d+)"(?P<attrs>[^>]*)>(?P<f><f(?:\s[^>]*)?>.*?</f>|<f(?:\s[^>]*)?/>)'
    r'(?:<v\s*/>|<v>\s*</v>|<v>[^<]*</v>)?</c>', re.S)


def _col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _xml_cell(m, val):
    attrs = re.sub(r'\s+t="[^"]*"', "", m.group("attrs"))
    f = m.group("f")
    ref = f'<c r="{m.group("col")}{m.group("row")}"'
    if val is None:
        return f"{ref}{attrs}>{f}</c>"
    if isinstance(val, XLErr):
        return f'{ref}{attrs} t="e">{f}<v>{html.escape(str(val))}</v></c>'
    if isinstance(val, bool):
        return f'{ref}{attrs} t="b">{f}<v>{1 if val else 0}</v></c>'
    if isinstance(val, (int, float)):
        x = float(val)
        if math.isnan(x) or math.isinf(x):
            return f'{ref}{attrs} t="e">{f}<v>#NUM!</v></c>'
        s = str(int(x)) if x == int(x) and abs(x) < 1e15 else repr(x)
        return f"{ref}{attrs}>{f}<v>{s}</v></c>"
    return f'{ref}{attrs} t="str">{f}<v>{html.escape(str(val), quote=False)}</v></c>'


def inject_cached_values(xlsx: bytes, values: dict) -> bytes:
    zin = zipfile.ZipFile(io.BytesIO(xlsx))
    wbxml = ET.fromstring(zin.read("xl/workbook.xml"))
    rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    rid_target = {r.get("Id"): r.get("Target") for r in rels}
    part_sheet = {}
    for s in wbxml.find("m:sheets", _NS):
        rid = s.get("{%s}id" % _NS["r"])
        tgt = rid_target.get(rid, "")
        tgt = tgt.lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        part_sheet[tgt] = s.get("name")
    by_sheet = {}
    for (sh, r, c), v in values.items():
        by_sheet.setdefault(sh, {})[(r, c)] = v
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            sh = part_sheet.get(item.filename)
            if sh and sh in by_sheet:
                vals = by_sheet[sh]
                text = data.decode("utf-8")

                def repl(m):
                    key = (int(m.group("row")), _col_index(m.group("col")))
                    if key not in vals:
                        return m.group(0)
                    return _xml_cell(m, vals[key])
                data = _CELL_F_RE.sub(repl, text).encode("utf-8")
            zout.writestr(item, data)
    return out.getvalue()


# --------------------------------------------------------------------------
# Ringkasan hasil untuk ditampilkan di aplikasi
# --------------------------------------------------------------------------
def make_summary(values: dict, base: dict, eng: Engine | None = None) -> dict:
    def v(sh, addr):
        col = re.match(r"[A-Z]+", addr).group(0)
        row = int(addr[len(col):])
        key = (sh, row, _col_index(col))
        if key in values:
            return values[key]
        return eng.cell(*key) if eng is not None else None   # sel input (bukan rumus)
    H = SH_HASIL
    s = {
        "ihsg_baseline": v(H, "D4"), "yield_baseline": v(H, "D5"),
        "pasar_escalated": v(H, "E8"), "pasar_severe": v(H, "F8"),
        "kredit": v(H, "E26"), "likuiditas_escalated": v(H, "E44"), "likuiditas_severe": v(H, "F44"),
        "pasar_suku_bunga_esc": v(SH_PASAR_ESC, "D20"), "pasar_saham_esc": v(SH_PASAR_ESC, "D27"),
        "pasar_suku_bunga_sev": v(SH_PASAR_SEV, "D20"), "pasar_saham_sev": v(SH_PASAR_SEV, "D27"),
        "pasar_pct_aset_esc": v(SH_PASAR_ESC, "E31"), "pasar_pct_aset_sev": v(SH_PASAR_SEV, "E31"),
        "kredit_non_mi": v(SH_CHK1, "H7"), "kredit_mi": v(SH_CHK1, "H8"),
        "kredit_pct_aset": v(SH_CHK1, "H13"),
        "lik_pct_esc": v(SH_LIK_ESC, "D17"), "lik_pct_sev": v(SH_LIK_SEV, "D17"),
        "total_1_6": v(SH_PASAR_ESC, "D15"),
    }
    rows = []
    labels = [("Total Aset Investasi", 10, 28, 46), ("SPI", 11, 29, 47), ("Laba Usaha", 12, 30, 48),
              ("Delta SPI", 13, 31, 49), ("Laba Usaha + Delta SPI", 14, 32, 50), ("ROI", 16, 34, 52)]
    for blok, idx in (("Risiko Pasar", 1), ("Risiko Kredit", 2), ("Risiko Likuiditas", 3)):
        for lab, *rr in labels:
            r = rr[idx - 1]
            rows.append({"Blok": blok, "Keterangan": lab,
                         "Existing": v(H, f"C{r}"), "Baseline": v(H, f"D{r}"),
                         "Escalated": v(H, f"E{r}"), "Severe": v(H, f"F{r}")})
    s["tabel"] = rows
    return s


# --------------------------------------------------------------------------
# Total instrumen terdampak suku bunga (= '1-Risiko Pasar'!D15), dihitung
# langsung dari RAW untuk mengisi otomatis "Nilai Wajar" kalkulator yield.
# --------------------------------------------------------------------------
def _add_months(d: dt.date, n: int) -> dt.date:
    y, m = d.year, d.month + n
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def total_terdampak_suku_bunga(raw: RawData, inp: StressInputs) -> float:
    """Miliar Rp. Meniru rumus baris 8–13 sheet 1-Risiko Pasar."""
    rep = month_end(inp.tahun, inp.bulan)
    tot = 0.0
    rd_all = rd_saham = 0.0
    for rec in raw.records():
        j = _jenis(rec)
        nw = _num(rec.get("nilaiwajar"))
        jt = rec.get("jatuhtempo")
        jt = jt.date() if isinstance(jt, dt.datetime) else jt
        has_dt = isinstance(jt, dt.date)
        if str(rec.get("type") or "").strip().lower() == "saham":
            rd_saham += nw
        if j == "REKSA":
            rd_all += nw
        if j == "TAB":
            tot += nw
        elif j in ("DEP", "DOC", "NCD", "SBI"):
            if has_dt:
                tot += nw
        elif "SBP" in j:
            if has_dt:
                tot += nw
        elif "OBL" in j or "SUK" in j:
            if has_dt:
                tot += nw
        if "EBA" in j and has_dt and "SBP" not in j:
            tot += nw
    posisi_saham_rd = sum(f.saham / 100 * f.nilai_wajar for f in inp.funds)
    tot += rd_all - rd_saham - posisi_saham_rd
    return tot / 1e9
