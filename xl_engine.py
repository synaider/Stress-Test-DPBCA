"""
xl_engine.py
============
Evaluator rumus Excel kecil (murni Python) untuk workbook openpyxl.

Dipakai untuk dua hal:
  1. Menghitung hasil stress test langsung di aplikasi (pratinjau sebelum
     file Excel dibuat).
  2. Mengisi "nilai tersimpan" (cached value) setiap sel rumus di file hasil,
     sehingga angka tetap tampil walaupun file dibuka di Protected View,
     pratinjau email, atau aplikasi yang tidak menghitung ulang rumus.

Fungsi yang didukung cukup untuk template Stress Test:
  SUM, SUMIF, SUMIFS, EDATE, ROUND, IF, IFERROR, ABS, MIN, MAX
Operator: + - * / ^ & % = <> < > <= >=, konstanta array {..}, referensi
antar-sheet, dan referensi kolom penuh (mis. 'RAW INVESTASI'!W:W).
"""
from __future__ import annotations

import calendar
import datetime as dt
import math
import re
from decimal import Decimal, ROUND_HALF_UP

from openpyxl.formula import Tokenizer
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.formula import ArrayFormula


# --------------------------------------------------------------------------
# Nilai error Excel
# --------------------------------------------------------------------------
class XLErr(str):
    """Nilai error Excel, mis. XLErr('#DIV/0!')."""


DIV0 = XLErr("#DIV/0!")
VALUE = XLErr("#VALUE!")
REF = XLErr("#REF!")
NAME = XLErr("#NAME?")
NUM = XLErr("#NUM!")
NA = XLErr("#N/A")


class _Abort(Exception):
    def __init__(self, err: XLErr):
        super().__init__(err)
        self.err = err


# --------------------------------------------------------------------------
# Tanggal <-> serial Excel (sistem 1900, termasuk "bug" 29 Feb 1900)
# --------------------------------------------------------------------------
_BASE_LATE = dt.date(1899, 12, 30)
_BASE_EARLY = dt.date(1899, 12, 31)


def date_to_serial(d) -> float:
    if isinstance(d, dt.datetime):
        frac = (d.hour * 3600 + d.minute * 60 + d.second + d.microsecond / 1e6) / 86400
        d0 = d.date()
    else:
        frac, d0 = 0.0, d
    if d0 >= dt.date(1900, 3, 1):
        return float((d0 - _BASE_LATE).days) + frac
    return float((d0 - _BASE_EARLY).days) + frac


def serial_to_date(s: float) -> dt.date:
    s = int(math.floor(s))
    if s >= 61:
        return _BASE_LATE + dt.timedelta(days=s)
    if s == 60:  # 29-02-1900 fiktif
        return dt.date(1900, 2, 28)
    return _BASE_EARLY + dt.timedelta(days=s)


def to_xl_value(v):
    """Ubah nilai Python dari openpyxl menjadi nilai 'Excel' (tanggal -> serial)."""
    if isinstance(v, (dt.datetime, dt.date)):
        return date_to_serial(v)
    if isinstance(v, dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second) / 86400
    return v


# --------------------------------------------------------------------------
# Referensi range
# --------------------------------------------------------------------------
_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_COL_RE = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW_RE = re.compile(r"^\$?(\d+)$")


class Ref:
    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2

    @property
    def shape(self):
        return (self.r2 - self.r1 + 1, self.c2 - self.c1 + 1)

    def key(self):
        return (self.sheet, self.r1, self.c1, self.r2, self.c2)


def _split_sheet(text: str, default_sheet: str):
    if "!" in text:
        sh, addr = text.rsplit("!", 1)
        if sh.startswith("'") and sh.endswith("'"):
            sh = sh[1:-1].replace("''", "'")
        return sh, addr
    return default_sheet, text


def _parse_part(p):
    m = _CELL_RE.match(p)
    if m:
        return int(m.group(2)), column_index_from_string(m.group(1).upper())
    m = _COL_RE.match(p)
    if m:
        return None, column_index_from_string(m.group(1).upper())
    m = _ROW_RE.match(p)
    if m:
        return int(m.group(1)), None
    raise ValueError(p)


# --------------------------------------------------------------------------
# Kriteria SUMIF/SUMIFS
# --------------------------------------------------------------------------
def _wild_to_regex(s: str):
    out, i = [], 0
    while i < len(s):
        ch = s[i]
        if ch == "~" and i + 1 < len(s) and s[i + 1] in "*?~":
            out.append(re.escape(s[i + 1]))
            i += 2
            continue
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    return re.compile("".join(out), re.IGNORECASE | re.DOTALL)


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def make_matcher(crit):
    """Bangun fungsi pencocok sesuai aturan kriteria Excel."""
    if isinstance(crit, XLErr):
        return lambda v: False
    if crit is None:
        crit = ""
    if isinstance(crit, bool):
        return lambda v, c=crit: isinstance(v, bool) and v == c
    if _is_num(crit):
        c = float(crit)
        return lambda v: _is_num(v) and float(v) == c
    s = str(crit)
    op = "="
    for o in (">=", "<=", "<>", "=", ">", "<"):
        if s.startswith(o):
            op, s = o, s[len(o):]
            break
    num = None
    try:
        if s.strip() != "":
            num = float(s)
    except ValueError:
        num = None
    if num is not None:
        if op == "=":
            return lambda v: _is_num(v) and float(v) == num
        if op == "<>":
            return lambda v: not (_is_num(v) and float(v) == num)
        cmp = {">": lambda a: a > num, "<": lambda a: a < num,
               ">=": lambda a: a >= num, "<=": lambda a: a <= num}[op]
        return lambda v: _is_num(v) and cmp(float(v))
    if s.upper() in ("TRUE", "FALSE") and op in ("=", "<>"):
        b = s.upper() == "TRUE"
        if op == "=":
            return lambda v: isinstance(v, bool) and v == b
        return lambda v: not (isinstance(v, bool) and v == b)
    # kriteria teks
    if op in ("=", "<>"):
        if s == "":
            if op == "=":
                return lambda v: v is None or v == ""
            return lambda v: not (v is None or v == "")
        rx = _wild_to_regex(s)
        if op == "=":
            return lambda v: isinstance(v, str) and rx.fullmatch(v) is not None
        return lambda v: not (isinstance(v, str) and rx.fullmatch(v) is not None)
    low = s.lower()
    cmp = {">": lambda a: a > low, "<": lambda a: a < low,
           ">=": lambda a: a >= low, "<=": lambda a: a <= low}[op]
    return lambda v: isinstance(v, str) and cmp(v.lower())


# --------------------------------------------------------------------------
# Parser rumus -> AST
# --------------------------------------------------------------------------
_PREC = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
         "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}
_UNARY_PREC = 7


class _Parser:
    def __init__(self, formula: str):
        toks = Tokenizer(formula).items
        self.t = [x for x in toks if x.type != "WHITE-SPACE"]
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else None

    def next(self):
        tok = self.peek()
        self.i += 1
        return tok

    def parse(self):
        if not self.t:
            return ("str", "")
        node = self.expr(0)
        if self.peek() is not None:
            raise ValueError("token sisa: %r" % (self.peek().value,))
        return node

    def expr(self, minprec):
        left = self.unary()
        while True:
            tok = self.peek()
            if tok is None:
                break
            if tok.type == "OPERATOR-POSTFIX":
                self.next()
                left = ("pct", left)
                continue
            if tok.type == "OPERATOR-INFIX" and tok.value in _PREC:
                p = _PREC[tok.value]
                if p < minprec:
                    break
                self.next()
                right = self.expr(p + 1)
                left = ("bin", tok.value, left, right)
                continue
            break
        return left

    def unary(self):
        tok = self.peek()
        if tok is not None and tok.type == "OPERATOR-PREFIX":
            self.next()
            operand = self.unary()
            # postfix % mengikat lebih kuat dari ^ tetapi lebih lemah dari negasi
            node = ("neg", operand) if tok.value == "-" else operand
            return node
        return self.primary()

    def primary(self):
        tok = self.next()
        if tok is None:
            raise ValueError("rumus tidak lengkap")
        if tok.type == "OPERAND":
            st = tok.subtype
            if st == "NUMBER":
                return ("num", float(tok.value))
            if st == "TEXT":
                return ("str", tok.value[1:-1].replace('""', '"'))
            if st == "LOGICAL":
                return ("bool", tok.value.upper() == "TRUE")
            if st == "ERROR":
                return ("err", XLErr(tok.value))
            if st == "RANGE":
                return ("ref", tok.value)
            raise ValueError("operand tidak dikenal %r" % tok.value)
        if tok.type == "FUNC" and tok.subtype == "OPEN":
            name = tok.value[:-1].upper()
            if name.startswith("_XLFN."):
                name = name[6:]
            args = []
            if self.peek() is not None and self.peek().type == "FUNC" and self.peek().subtype == "CLOSE":
                self.next()
                return ("func", name, args)
            while True:
                nxt = self.peek()
                if nxt is not None and nxt.type == "SEP" and nxt.subtype == "ARG":
                    args.append(("blank",))
                    self.next()
                    continue
                args.append(self.expr(0))
                nxt = self.next()
                if nxt is None:
                    raise ValueError("kurung fungsi tidak ditutup")
                if nxt.type == "SEP" and nxt.subtype == "ARG":
                    continue
                if nxt.type == "FUNC" and nxt.subtype == "CLOSE":
                    break
                raise ValueError("token tak terduga %r" % nxt.value)
            return ("func", name, args)
        if tok.type == "PAREN" and tok.subtype == "OPEN":
            node = self.expr(0)
            close = self.next()
            if close is None or close.type != "PAREN":
                raise ValueError("kurung tidak ditutup")
            return node
        if tok.type == "ARRAY" and tok.subtype == "OPEN":
            rows, row = [], []
            while True:
                nxt = self.peek()
                if nxt is None:
                    raise ValueError("array tidak ditutup")
                if nxt.type == "ARRAY" and nxt.subtype == "CLOSE":
                    self.next()
                    rows.append(row)
                    break
                if nxt.type == "SEP":
                    self.next()
                    if nxt.subtype == "ROW":
                        rows.append(row)
                        row = []
                    continue
                el = self.unary()
                row.append(el)
            return ("arr", rows)
        raise ValueError("token tak terduga %r (%s)" % (tok.value, tok.type))


# --------------------------------------------------------------------------
# Mesin evaluasi
# --------------------------------------------------------------------------
class Engine:
    def __init__(self, wb):
        self.wb = wb
        self.values = {}
        self._stack = set()
        self._ast_cache = {}
        self._col_cache = {}
        self._max_row = {ws.title: ws.max_row for ws in wb.worksheets}

    # ---------------- akses sel ----------------
    def is_formula_cell(self, cell):
        v = cell.value
        if isinstance(v, ArrayFormula):
            return True
        return isinstance(v, str) and cell.data_type == "f"

    def cell(self, sheet, row, col):
        key = (sheet, row, col)
        if key in self.values:
            return self.values[key]
        try:
            ws = self.wb[sheet]
        except KeyError:
            return REF
        c = ws._cells.get((row, col))
        if c is None:
            return None
        v = c.value
        if isinstance(v, ArrayFormula):
            formula = v.text
        elif isinstance(v, str) and c.data_type == "f":
            formula = v
        else:
            val = to_xl_value(v)
            self.values[key] = val
            return val
        if key in self._stack:
            return REF  # referensi melingkar
        self._stack.add(key)
        try:
            val = self.eval_formula(formula, sheet)
        finally:
            self._stack.discard(key)
        self.values[key] = val
        return val

    def eval_formula(self, formula: str, sheet: str):
        f = formula if formula.startswith("=") else "=" + formula
        ast = self._ast_cache.get(f)
        if ast is None:
            try:
                ast = _Parser(f).parse()
            except Exception:
                ast = ("err", NAME)
            self._ast_cache[f] = ast
        try:
            val = self.ev(ast, sheet)
        except _Abort as a:
            return a.err
        except ZeroDivisionError:
            return DIV0
        except (OverflowError, ValueError):
            return NUM
        val = self.scalarize(val)
        return val

    def scalarize(self, val):
        if isinstance(val, Ref):
            v = self.cell(val.sheet, val.r1, val.c1)
            return 0.0 if v is None else v
        if isinstance(val, list):
            if not val or not val[0]:
                return VALUE
            return self.scalarize(val[0][0])
        if val is None:
            return 0.0
        return val

    # ---------------- referensi ----------------
    def make_ref(self, text, sheet):
        sh, addr = _split_sheet(text, sheet)
        if sh not in self._max_row:
            raise _Abort(REF)
        parts = addr.split(":")
        try:
            a = _parse_part(parts[0])
            b = _parse_part(parts[1]) if len(parts) > 1 else a
        except ValueError:
            raise _Abort(NAME)
        r1, c1 = a
        r2, c2 = b
        maxr = max(self._max_row[sh], 1)
        if r1 is None or r2 is None:  # kolom penuh
            r1, r2 = 1, maxr
        if c1 is None or c2 is None:  # baris penuh
            c1, c2 = 1, max(self.wb[sh].max_column, 1)
        if r1 > r2:
            r1, r2 = r2, r1
        if c1 > c2:
            c1, c2 = c2, c1
        return Ref(sh, r1, c1, r2, c2)

    def ref_values(self, ref: Ref):
        """Daftar nilai (baris-mayor) untuk sebuah range, di-cache."""
        k = ref.key()
        vals = self._col_cache.get(k)
        if vals is None:
            vals = [self.cell(ref.sheet, r, c)
                    for r in range(ref.r1, ref.r2 + 1)
                    for c in range(ref.c1, ref.c2 + 1)]
            self._col_cache[k] = vals
        return vals

    # ---------------- evaluasi AST ----------------
    def ev(self, node, sheet):
        kind = node[0]
        if kind == "num":
            return node[1]
        if kind == "str":
            return node[1]
        if kind == "bool":
            return node[1]
        if kind == "err":
            return node[1]
        if kind == "blank":
            return None
        if kind == "ref":
            return self.make_ref(node[1], sheet)
        if kind == "arr":
            return [[self.scalarize(self.ev(e, sheet)) for e in row] for row in node[1]]
        if kind == "neg":
            return self._map1(self.ev(node[1], sheet), lambda x: -self.num(x))
        if kind == "pct":
            return self._map1(self.ev(node[1], sheet), lambda x: self.num(x) / 100.0)
        if kind == "bin":
            return self._binop(node[1], self.ev(node[2], sheet), self.ev(node[3], sheet))
        if kind == "func":
            return self._call(node[1], node[2], sheet)
        raise _Abort(NAME)

    def _map1(self, v, fn):
        if isinstance(v, Ref):
            if v.shape == (1, 1):
                v = self.cell(v.sheet, v.r1, v.c1)
            else:
                raise _Abort(VALUE)
        if isinstance(v, list):
            return [[self._safe(fn, x) for x in row] for row in v]
        return fn(v)

    def _safe(self, fn, x):
        try:
            return fn(x)
        except _Abort as a:
            return a.err
        except ZeroDivisionError:
            return DIV0

    def num(self, v):
        if isinstance(v, Ref):
            if v.shape != (1, 1):
                raise _Abort(VALUE)
            v = self.cell(v.sheet, v.r1, v.c1)
        if isinstance(v, XLErr):
            raise _Abort(v)
        if v is None:
            return 0.0
        if isinstance(v, bool):
            return 1.0 if v else 0.0
        if _is_num(v):
            return float(v)
        if isinstance(v, str):
            s = v.strip()
            if s == "":
                raise _Abort(VALUE)
            try:
                if s.endswith("%"):
                    return float(s[:-1]) / 100.0
                return float(s)
            except ValueError:
                raise _Abort(VALUE)
        raise _Abort(VALUE)

    def text(self, v):
        if isinstance(v, Ref):
            if v.shape != (1, 1):
                raise _Abort(VALUE)
            v = self.cell(v.sheet, v.r1, v.c1)
        if isinstance(v, XLErr):
            raise _Abort(v)
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if _is_num(v):
            return fmt_general(v)
        return str(v)

    def _binop(self, op, a, b):
        if isinstance(a, list) or isinstance(b, list):
            a = self._as_array(a)
            b = self._as_array(b)
            rows = max(len(a), len(b))
            cols = max(len(a[0]), len(b[0]))
            out = []
            for i in range(rows):
                row = []
                for j in range(cols):
                    x = a[i if len(a) > 1 else 0][j if len(a[0]) > 1 else 0]
                    y = b[i if len(b) > 1 else 0][j if len(b[0]) > 1 else 0]
                    try:
                        row.append(self._scalar_op(op, x, y))
                    except _Abort as e:
                        row.append(e.err)
                    except ZeroDivisionError:
                        row.append(DIV0)
                out.append(row)
            return out
        return self._scalar_op(op, a, b)

    def _as_array(self, v):
        if isinstance(v, list):
            return v
        if isinstance(v, Ref):
            vals = self.ref_values(v)
            h, w = v.shape
            return [vals[i * w:(i + 1) * w] for i in range(h)]
        return [[v]]

    def _scalar_op(self, op, a, b):
        if op == "&":
            return self.text(a) + self.text(b)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            return self._compare(op, a, b)
        x, y = self.num(a), self.num(b)
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise _Abort(DIV0)
            return x / y
        if op == "^":
            if x < 0 and y != int(y):
                raise _Abort(NUM)
            if x == 0 and y < 0:
                raise _Abort(DIV0)
            return float(x) ** float(y)
        raise _Abort(NAME)

    def _compare(self, op, a, b):
        def deref(v):
            if isinstance(v, Ref):
                v = self.cell(v.sheet, v.r1, v.c1)
            if isinstance(v, XLErr):
                raise _Abort(v)
            return v
        a, b = deref(a), deref(b)

        def rank(v):
            if v is None:
                return (0, 0)
            if isinstance(v, bool):
                return (3, int(v))
            if _is_num(v):
                return (1, float(v))
            return (2, str(v).lower())
        if a is None:
            a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
        if b is None:
            b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)
        ra, rb = rank(a), rank(b)
        res = {"=": ra == rb, "<>": ra != rb, "<": ra < rb,
               ">": ra > rb, "<=": ra <= rb, ">=": ra >= rb}[op]
        return res

    # ---------------- fungsi ----------------
    def _call(self, name, args, sheet):
        fn = getattr(self, "f_" + name, None)
        if fn is None:
            raise _Abort(NAME)
        if name in ("IF", "IFERROR"):
            return fn(args, sheet)
        # Error pada argumen menjadi NILAI (seperti Excel), bukan menghentikan
        # seluruh rumus. Contoh: SUMIFS(..., ">" & EDATE(teks,3)) -> kriteria
        # bernilai #VALUE! sehingga tidak ada sel yang cocok -> hasil 0.
        vals = []
        for a in args:
            try:
                vals.append(self.ev(a, sheet))
            except _Abort as e:
                vals.append(e.err)
            except ZeroDivisionError:
                vals.append(DIV0)
        return fn(*vals)

    def _iter_numbers(self, v):
        if isinstance(v, Ref):
            for x in self.ref_values(v):
                if isinstance(x, XLErr):
                    raise _Abort(x)
                if _is_num(x):
                    yield float(x)
        elif isinstance(v, list):
            for row in v:
                for x in row:
                    if isinstance(x, XLErr):
                        raise _Abort(x)
                    if _is_num(x):
                        yield float(x)
        elif v is None:
            return
        else:
            yield self.num(v)

    def f_SUM(self, *args):
        return math.fsum(x for a in args for x in self._iter_numbers(a))

    def f_MIN(self, *args):
        xs = [x for a in args for x in self._iter_numbers(a)]
        return min(xs) if xs else 0.0

    def f_MAX(self, *args):
        xs = [x for a in args for x in self._iter_numbers(a)]
        return max(xs) if xs else 0.0

    def f_ABS(self, x):
        return self._map1(x, lambda v: abs(self.num(v)))

    def f_ROUND(self, x, n=0.0):
        digits = int(self.num(n))

        def rnd(v):
            q = Decimal(1).scaleb(-digits)
            return float(Decimal(repr(self.num(v))).quantize(q, rounding=ROUND_HALF_UP))
        return self._map1(x, rnd)

    def f_IF(self, args, sheet):
        cond = self.scalarize(self.ev(args[0], sheet))
        if isinstance(cond, XLErr):
            return cond
        truthy = bool(self.num(cond)) if not isinstance(cond, str) else bool(cond)
        if truthy:
            return self.ev(args[1], sheet) if len(args) > 1 else True
        return self.ev(args[2], sheet) if len(args) > 2 else False

    def f_IFERROR(self, args, sheet):
        try:
            v = self.scalarize(self.ev(args[0], sheet))
        except _Abort as a:
            v = a.err
        except ZeroDivisionError:
            v = DIV0
        if isinstance(v, XLErr):
            return self.ev(args[1], sheet)
        return v

    def f_EDATE(self, start, months):
        def one(s, m):
            s = self.num(s)
            m = int(self.num(m))
            d = serial_to_date(s)
            y, mo = d.year, d.month + m
            y += (mo - 1) // 12
            mo = (mo - 1) % 12 + 1
            if y < 1900 and not (y == 1899 and mo == 12):
                raise _Abort(NUM)
            day = min(d.day, calendar.monthrange(y, mo)[1])
            return date_to_serial(dt.date(y, mo, day))
        if isinstance(start, Ref) and start.shape == (1, 1):
            start = self.cell(start.sheet, start.r1, start.c1)
        if isinstance(start, str):
            raise _Abort(VALUE)
        return one(start, months)

    def _crit_list(self, crit):
        """Kriteria bisa skalar atau array; kembalikan (bentuk, daftar)."""
        if isinstance(crit, Ref):
            if crit.shape == (1, 1):
                return None, [self.cell(crit.sheet, crit.r1, crit.c1)]
            vals = self.ref_values(crit)
            h, w = crit.shape
            return (h, w), vals
        if isinstance(crit, list):
            h, w = len(crit), len(crit[0]) if crit else 0
            return (h, w), [x for row in crit for x in row]
        return None, [crit]

    def f_SUMIF(self, rng, crit, sum_rng=None):
        if not isinstance(rng, Ref):
            raise _Abort(VALUE)
        if sum_rng is None:
            sum_rng = rng
        if not isinstance(sum_rng, Ref):
            raise _Abort(VALUE)
        h, w = rng.shape
        sum_rng = Ref(sum_rng.sheet, sum_rng.r1, sum_rng.c1,
                      sum_rng.r1 + h - 1, sum_rng.c1 + w - 1)
        return self.f_SUMIFS(sum_rng, rng, crit)

    def f_SUMIFS(self, sum_rng, *pairs):
        if not isinstance(sum_rng, Ref) or len(pairs) % 2:
            raise _Abort(VALUE)
        svals = self.ref_values(sum_rng)
        ranges, crits, shape = [], [], None
        for i in range(0, len(pairs), 2):
            rg, cr = pairs[i], pairs[i + 1]
            if not isinstance(rg, Ref) or rg.shape != sum_rng.shape:
                raise _Abort(VALUE)
            shp, lst = self._crit_list(cr)
            if shp is not None:
                shape = shp
            ranges.append(self.ref_values(rg))
            crits.append(lst)
        n = max(len(c) for c in crits)

        def total(k):
            matchers = [make_matcher(c[k if len(c) > 1 else 0]) for c in crits]
            s = 0.0
            for idx, sv in enumerate(svals):
                if not _is_num(sv):
                    if isinstance(sv, XLErr) and all(m(r[idx]) for m, r in zip(matchers, ranges)):
                        raise _Abort(sv)
                    continue
                ok = True
                for m, r in zip(matchers, ranges):
                    if not m(r[idx]):
                        ok = False
                        break
                if ok:
                    s += float(sv)
            return s
        if shape is None:
            return total(0)
        res = [total(k) for k in range(n)]
        h, w = shape
        if h * w != n:
            h, w = 1, n
        return [res[i * w:(i + 1) * w] for i in range(h)]

    # GETPIVOTDATA tidak dievaluasi (template hasil tidak lagi memakainya)
    def f_GETPIVOTDATA(self, *args):
        raise _Abort(REF)

    # ---------------- hitung semua ----------------
    def evaluate_all(self):
        """Hitung semua sel rumus. Kembalikan {(sheet,row,col): nilai}."""
        out = {}
        for ws in self.wb.worksheets:
            for (r, c), cell in list(ws._cells.items()):
                if self.is_formula_cell(cell):
                    out[(ws.title, r, c)] = self.cell(ws.title, r, c)
        return out


def fmt_general(v: float) -> str:
    """Format angka seperti 'General' Excel saat digabung teks (&)."""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    f = float(v)
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    s = "%.15g" % f
    return s
