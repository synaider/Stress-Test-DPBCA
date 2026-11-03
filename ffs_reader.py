"""
ffs_reader.py
=============
Mengambil dan membaca Fund Fact Sheet (FFS) reksa dana untuk sheet
"Komposisi RD": persentase Saham, Obligasi, dan Pasar Uang (Kas/Lainnya).

Alur:
  1. Katalog reksa dana (config/katalog_reksadana.json) berisi daftar link FFS.
     Link boleh memakai penanda bulan, contoh:
         {bulan} {yyyy}   -> September 2026   (bulan data FFS = bulan laporan)
         {pub_yyyy}-{pub_m} -> 2026-10        (bulan terbit = bulan laporan + 1)
  2. fetch_latest_ffs() mencoba setiap link, mengunduh PDF, membaca alokasi
     aset, lalu memilih FFS yang tanggal datanya sama dengan bulan laporan.
  3. Bila semua link gagal (situs berubah / tidak ada internet), pengguna bisa
     mengunggah PDF FFS sendiri -> parse_ffs_pdf(), atau mengetik manual.
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CATALOG_PATH = BASE_DIR / "config" / "katalog_reksadana.json"

BULAN_ID = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
            "September", "Oktober", "November", "Desember"]
BULAN_EN = ["", "January", "February", "March", "April", "May", "June", "July", "August",
            "September", "October", "November", "December"]
_MONTH_WORDS = {}
for _i in range(1, 13):
    for _w in (BULAN_ID[_i], BULAN_EN[_i], BULAN_ID[_i][:3], BULAN_EN[_i][:3]):
        _MONTH_WORDS[_w.lower()] = _i
_MONTH_WORDS.update({"sept": 9, "agu": 8, "agt": 8, "ags": 8, "okt": 10, "des": 12, "mei": 5,
                     "juni": 6, "juli": 7, "nop": 11, "pebruari": 2})

DEFAULT_CATALOG = [
    {
        "nama": "Schroder Dana Terpadu II",
        "mi": "PT Schroder Investment Management Indonesia",
        "jenis": "Campuran",
        "urls": [
            "https://www.smbci.com/pdf/sinaya/investment-product/{pub_yyyy}-{pub_m}/ffs----schroder--dana--terpadu--ii.pdf",
            "https://static-invest.cermati.com/mutualfund/documents/fund-fact-sheet/Schroder_Investment_Management_Indonesia_Schroder_Dana_Terpadu_II_{bulan}_{yyyy}.pdf",
            "https://www.dbs.com/id/iwov-resources/flp/pdf/FFS-Prospectus/reksadana-campuran/ffs/SCHRODER_DANA_TERPADU_II_EN.pdf",
            "https://www.maybank.co.id/premier/aset-enhancer/-/media/Downloaded Content/Investment/Fund-Fact-Sheet/Schroders/FFS-Schroder-Dana-Terpadu-II.pdf",
        ],
        "halaman": [
            "https://www.bareksa.com/id/data/reksadana/307/schroder-dana-terpadu-ii",
            "https://www.bca.co.id/id/Individu/produk/Investasi-dan-Asuransi/reksadana/Pilihan-Produk-Reksadana/Schroder-Dana-Terpadu-II",
        ],
    },
    {
        "nama": "Trimegah Balanced Absolute Strategy Kelas A",
        "mi": "PT Trimegah Asset Management",
        "jenis": "Campuran",
        "urls": [
            "https://cache.makmur.id/files/products/60d1e69d4cc7bb3913edb067/factsheet.pdf",
            "https://static-invest.cermati.com/mutualfund/documents/fund-fact-sheet/Trimegah_Asset_Management_Trimegah_Balanced_Absolute_Strategy_Kelas_A_{bulan}_{yyyy}.pdf",
            "https://media.bareksa.com/uploads/file_doc/{pub_yyyy}/{pub_mm}/ACTBKAS_factsheet.pdf",
        ],
        "halaman": [
            "https://www.trimegah-am.com/product/view/reksa-dana-campuran/trimegah-balanced-absolute-strategy",
            "https://invest.cermati.com/reksadana/campuran/trimegah-balanced-absolute-strategy-kelas-a-lg002mxc01bast01",
        ],
    },
    {
        "nama": "Trimegah Balanced Absolute Strategy Low Volatility",
        "mi": "PT Trimegah Asset Management",
        "jenis": "Campuran",
        "urls": [
            "https://cache.makmur.id/files/products/65b070ed3bd47139ba4debf0/factsheet.pdf",
            "https://static-invest.cermati.com/mutualfund/documents/fund-fact-sheet/Trimegah_Asset_Management_Trimegah_Balanced_Absolute_Strategy_Low_Volatility_{bulan}_{yyyy}.pdf",
        ],
        "halaman": ["https://www.trimegah-am.com/"],
    },
]


# --------------------------------------------------------------------------
# Katalog
# --------------------------------------------------------------------------
def load_catalog() -> list:
    try:
        data = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
        if isinstance(data, list) and data:
            return data
    except Exception:
        pass
    return [dict(x) for x in DEFAULT_CATALOG]


def save_catalog(catalog: list):
    CATALOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CATALOG_PATH.write_text(json.dumps(catalog, ensure_ascii=False, indent=2), encoding="utf-8")


def _norm(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def find_in_catalog(catalog: list, nama: str):
    k = _norm(nama)
    for f in catalog:
        if _norm(f.get("nama")) == k:
            return f
    return None


def _title_part(s: str) -> str:
    words = []
    for w in re.split(r"\s+", str(s).strip()):
        if not w:
            continue
        if w.upper() in ("PT", "TBK"):
            if w.upper() == "PT":
                continue
            words.append("Tbk")
            continue
        if re.fullmatch(r"[IVX]+", w.upper()):
            words.append(w.upper())
        else:
            words.append(w[:1].upper() + w[1:].lower())
    return "_".join(re.sub(r"[^A-Za-z0-9_]", "", x) for x in words)


def generic_entry(nama: str, mi: str = "") -> dict:
    """Entri katalog otomatis untuk reksa dana yang belum ada di katalog."""
    urls = []
    if mi:
        urls.append("https://static-invest.cermati.com/mutualfund/documents/fund-fact-sheet/"
                    f"{_title_part(mi)}_{_title_part(nama)}_{{bulan}}_{{yyyy}}.pdf")
    return {"nama": nama, "mi": mi, "jenis": "", "urls": urls, "halaman": []}


def search_links(nama: str, year: int, month: int) -> list:
    from urllib.parse import quote_plus
    q = quote_plus(f"fund fact sheet {nama} {BULAN_ID[month]} {year} pdf")
    q2 = quote_plus(f"{nama} fund fact sheet")
    return [("Cari di Google", f"https://www.google.com/search?q={q}"),
            ("Cari di Bareksa", f"https://www.bareksa.com/cari?search={quote_plus(nama)}"),
            ("Cari di Bing", f"https://www.bing.com/search?q={q2}")]


def expand_url(tpl: str, year: int, month: int) -> str:
    py, pm = (year + 1, 1) if month == 12 else (year, month + 1)
    rep = {"{yyyy}": str(year), "{mm}": f"{month:02d}", "{m}": str(month),
           "{bulan}": BULAN_ID[month], "{month}": BULAN_EN[month],
           "{pub_yyyy}": str(py), "{pub_mm}": f"{pm:02d}", "{pub_m}": str(pm),
           "{pub_bulan}": BULAN_ID[pm]}
    out = tpl
    for k, v in rep.items():
        out = out.replace(k, v)
    return out.replace(" ", "%20")


# --------------------------------------------------------------------------
# Membaca PDF
# --------------------------------------------------------------------------
def pdf_text(data: bytes) -> str:
    texts = []
    try:
        import pdfplumber  # lebih rapi untuk tata letak kolom
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for p in pdf.pages[:4]:
                texts.append(p.extract_text() or "")
    except Exception:
        texts = []
    if not "".join(texts).strip():
        try:
            from pypdf import PdfReader
            rd = PdfReader(io.BytesIO(data))
            texts = [(pg.extract_text() or "") for pg in rd.pages[:4]]
        except Exception as e:
            raise RuntimeError(f"PDF tidak bisa dibaca: {e}")
    return "\n".join(texts)


SAHAM_LBL = r"(?:saham|ekuitas|equities|equity|stocks?|shares)"
OBL_LBL = (r"(?:obligasi|efek\s+bersifat\s+utang|efek\s+utang|surat\s+utang|sukuk|debt\s+securities|"
           r"bonds?|fixed\s+income)")
PU_LBL = (r"(?:instrumen\s+pasar\s+uang|pasar\s+uang|money\s+market|kas\s*(?:&|dan)\s*setara\s+kas|kas|cash"
          r"(?:\s*&\s*equivalents?)?|deposito|lainnya|lain-lain|others?)")
SECTION_RE = re.compile(r"(alokasi\s+aset|asset\s+allocation|komposisi\s+portofolio|alokasi\s+portofolio|"
                        r"komposisi\s+aset|portfolio\s+allocation|komposisi\s+investasi)", re.I)
PCT = r"(\d{1,3}(?:[.,]\d{1,2})?)\s*%"
NOT_RANGE = r"(?!\s*[-–]\s*\d)"


def _f(s: str) -> float:
    return float(s.replace(",", "."))


@dataclass
class Allocation:
    saham: float = 0.0
    obligasi: float = 0.0
    pasar_uang: float = 0.0
    as_of: tuple | None = None       # (tahun, bulan)
    metode: str = ""
    catatan: list = field(default_factory=list)

    @property
    def total(self):
        return self.saham + self.obligasi + self.pasar_uang


def _pairs_adjacent(win: str):
    res = {}
    for key, lbl in (("saham", SAHAM_LBL), ("obligasi", OBL_LBL), ("pasar_uang", PU_LBL)):
        rx = re.compile(lbl + r"\s*[:=]?\s*" + PCT + NOT_RANGE, re.I)
        vals = [_f(m.group(1)) for m in rx.finditer(win)]
        if vals:
            res[key] = vals
    return res


def _pairs_ordered(win: str):
    """Tata letak: daftar label lalu daftar nilai (atau sebaliknya), dipasangkan berurutan."""
    lab_rx = re.compile(r"\b(" + SAHAM_LBL + r"|" + OBL_LBL + r"|" + PU_LBL + r")\b", re.I)
    labels = []
    for m in lab_rx.finditer(win):
        w = m.group(1).lower()
        if re.fullmatch(SAHAM_LBL, w, re.I):
            k = "saham"
        elif re.fullmatch(OBL_LBL, w, re.I):
            k = "obligasi"
        else:
            k = "pasar_uang"
        labels.append(k)
    vals = [_f(m.group(1)) for m in re.finditer(r"(\d{1,3}[.,]\d{1,2})\s*%" + NOT_RANGE, win)]
    return labels, vals


def parse_allocation(text: str) -> Allocation | None:
    t = text.replace(" ", " ")
    windows = []
    for m in SECTION_RE.finditer(t):
        windows.append(t[m.end(): m.end() + 700])
    windows.append(t)  # fallback: seluruh teks
    best = None
    for idx, win in enumerate(windows):
        # Metode A: label langsung diikuti angka
        adj = _pairs_adjacent(win)
        if "saham" in adj and ("obligasi" in adj or "pasar_uang" in adj):
            a = Allocation(saham=adj["saham"][0], obligasi=adj.get("obligasi", [0])[0],
                           pasar_uang=adj.get("pasar_uang", [0])[0], metode="label-angka")
            if 95 <= a.total <= 105:
                return a
            best = best or a
        # Metode B: urutan label & angka
        if idx < len(windows) - 1:
            labels, vals = _pairs_ordered(win)
            seen, labs = set(), []
            for k in labels:
                if k not in seen or k == "pasar_uang":
                    labs.append(k)
                    seen.add(k)
            n = min(len(labs), len(vals))
            for start in range(0, max(1, len(vals) - n + 1)):
                cand = Allocation(metode="urutan")
                for k, v in zip(labs[:n], vals[start:start + n]):
                    setattr(cand, k, getattr(cand, k) + v)
                if cand.saham and 95 <= cand.total <= 105:
                    return cand
    return best


DATE_RES = [
    re.compile(r"(\d{1,2})[\s\-/]+([A-Za-z]{3,9})[\s\-/,]+(20\d{2})"),
    re.compile(r"\b([A-Za-z]{3,9})[\s\-/,]+(20\d{2})\b"),
    re.compile(r"\b(\d{1,2})/(\d{1,2})/(20\d{2})\b"),
]


def parse_as_of(text: str):
    t = text.replace(" ", " ")
    # prioritaskan kalimat "as of" / "per tanggal" / "data per"
    for m in re.finditer(r"(as\s+of|per\s+tanggal|data\s+per|posisi|per)\s+(\d{1,2}[\s\-/]+[A-Za-z]{3,9}[\s\-/]+20\d{2})",
                         t, re.I):
        d = _date_from(m.group(2))
        if d:
            return d
    def skip(pos):  # abaikan tanggal terbit / tanggal efektif
        before = t[max(0, pos - 40):pos].lower()
        return any(w in before for w in ("issue date", "terbit", "diterbitkan", "efektif", "effective",
                                         "peluncuran", "launch", "inception"))
    found = []
    for m in DATE_RES[0].finditer(t):
        mo = _MONTH_WORDS.get(m.group(2).lower())
        if mo and not skip(m.start()):
            found.append((int(m.group(3)), mo))
    if not found:
        for m in DATE_RES[1].finditer(t):
            mo = _MONTH_WORDS.get(m.group(1).lower())
            if mo and not skip(m.start()):
                found.append((int(m.group(2)), mo))
    # FFS memuat tanggal data terakhir sebagai tanggal paling akhir di dokumen
    return max(found) if found else None


def _date_from(s: str):
    m = DATE_RES[0].search(s)
    if m:
        mo = _MONTH_WORDS.get(m.group(2).lower())
        if mo:
            return (int(m.group(3)), mo)
    return None


def parse_ffs_pdf(data: bytes) -> tuple:
    """Kembalikan (Allocation|None, teks_pdf)."""
    text = pdf_text(data)
    alloc = parse_allocation(text)
    if alloc is not None:
        alloc.as_of = parse_as_of(text)
    return alloc, text


# --------------------------------------------------------------------------
# Unduh otomatis
# --------------------------------------------------------------------------
@dataclass
class FetchResult:
    ok: bool
    alokasi: Allocation | None = None
    url: str = ""
    pdf: bytes = b""
    log: list = field(default_factory=list)


HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36"),
    "Accept": "application/pdf,*/*;q=0.8",
}


def _download(url: str, timeout=20) -> bytes:
    import requests
    r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
    r.raise_for_status()
    if not r.content.startswith(b"%PDF"):
        raise ValueError("bukan file PDF")
    return r.content


def fetch_latest_ffs(entry: dict, year: int, month: int, extra_urls=None, downloader=_download) -> FetchResult:
    """Coba semua link FFS di katalog. Utamakan FFS yang bulan datanya = bulan laporan."""
    urls = [expand_url(u, year, month) for u in (list(extra_urls or []) + list(entry.get("urls", [])))]
    log, fallback = [], None
    for url in dict.fromkeys(urls):
        try:
            data = downloader(url)
        except Exception as e:
            log.append(f"✗ {url} — {type(e).__name__}: {str(e)[:80]}")
            continue
        try:
            alloc, _ = parse_ffs_pdf(data)
        except Exception as e:
            log.append(f"✗ {url} — PDF tidak terbaca ({e})")
            continue
        if alloc is None:
            log.append(f"✗ {url} — alokasi aset tidak ditemukan di PDF")
            continue
        tag = f"{BULAN_ID[alloc.as_of[1]]} {alloc.as_of[0]}" if alloc.as_of else "tanggal tidak terbaca"
        if alloc.as_of == (year, month):
            log.append(f"✓ {url} — FFS {tag}")
            return FetchResult(True, alloc, url, data, log)
        log.append(f"• {url} — FFS {tag} (bukan bulan laporan)")
        if fallback is None or (alloc.as_of and fallback[0].as_of and alloc.as_of > fallback[0].as_of):
            fallback = (alloc, url, data)
    if fallback:
        return FetchResult(True, fallback[0], fallback[1], fallback[2], log)
    return FetchResult(False, None, "", b"", log)
