"""Karobka VBA modullarining (cheklangan VBA to'plami) Python'ga o'giruvchisi.

Maqsad: src/*.bas dagi geometriyani Corelsiz ishga tushirib, chizmalarni tekshirish (tools/preview.py).
Faqat shu loyihada ishlatilgan VBA qurilmalari qo'llanadi. Noma'lum narsa uchrasa xato beradi.
"""
import keyword
import math
import re

# ------------------------------------------------------------------ ish vaqti yordamchilari


class VStr(str):
    """VBA '&' operatori: str | x"""

    def __or__(self, o):
        return VStr(str(self) + fmt(o))

    def __ror__(self, o):
        return VStr(fmt(o) + str(self))


def fmt(v):
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


class VArr:
    def __init__(self, lo=0, hi=-1):
        self.lo, self.a = lo, [0] * (hi - lo + 1)

    def redim(self, lo, hi, preserve=False):
        old = self.a
        self.lo, self.a = int(lo), [0] * (int(hi) - int(lo) + 1)
        if preserve:
            self.a[:len(old)] = old[:len(self.a)]

    def _i(self, i):
        i = int(round(i)) - self.lo
        if i < 0 or i >= len(self.a):
            raise IndexError("massiv chegarasidan tashqari: %d" % (i + self.lo))
        return i

    def __call__(self, i):
        return self.a[self._i(i)]

    def __setitem__(self, i, v):
        self.a[self._i(i)] = v

    def ubound(self):
        return self.lo + len(self.a) - 1


class VDict:
    def __init__(self):
        self.d = {}

    def __call__(self, k):
        return self.d[k]

    def __setitem__(self, k, v):
        self.d[k] = v

    def exists(self, k):
        return k in self.d

    def add(self, k, v):
        if k in self.d:
            raise KeyError("takroriy kalit " + k)
        self.d[k] = v


class PArr(list):
    """ParamArray: v(i) ko'rinishida murojaat"""

    def __call__(self, i):
        return self[int(round(i))]


def _rng(a, b, s=1):
    i = a
    while (s > 0 and i <= b + 1e-12) or (s < 0 and i >= b - 1e-12):
        yield i
        i += s


RUNTIME = dict(
    PArr=PArr, VStr=VStr, VArr=VArr, VDict=VDict, _rng=_rng, math=math,
    ubound=lambda a: a.ubound() if isinstance(a, VArr) else len(a) - 1,
    cos=math.cos, sin=math.sin, tan=math.tan,
    instr=lambda a, b: a.find(b) + 1, sqr=math.sqrt, abs=abs, atn=math.atan,
    int=lambda v: math.floor(v), fix=lambda v: math.trunc(v),
    round=lambda v, n=0: round(v, int(n)),
    cstr=lambda v: VStr(fmt(v)), val=lambda s: float(re.match(r"\s*([-+]?\d*\.?\d*)", s).group(1) or 0),
    replace=lambda s, a, b: VStr(s.replace(a, b)), split=lambda s, d: PArr(VStr(x) for x in s.split(d)),
    len=lambda s: len(s), createobject=lambda name: VDict(),
)

# ------------------------------------------------------------------ o'giruvchi

TOK = re.compile(r'"(?:[^"]|"")*"|\d+\.\d*(?:[eE][-+]?\d+)?|\d*\.\d+(?:[eE][-+]?\d+)?|\d+(?:[eE][-+]?\d+)?'
                 r'|[A-Za-z_]\w*|<>|<=|>=|:=|[-+*/\\^&=<>(),.:]|\s+')
PYKW = set(keyword.kwlist) | {"print", "exec", "type", "id", "list", "dict", "str", "min", "max", "sum", "map", "range"}
OPS = {"and": " and ", "or": " or ", "not": " not ", "mod": " % ", "true": "True", "false": "False"}


def tokens(s):
    out, pos = [], 0
    while pos < len(s):
        m = TOK.match(s, pos)
        if not m:
            raise SyntaxError("tushunarsiz belgi: %r" % s[pos:])
        out.append(m.group(0))
        pos = m.end()
    return out


def strip_comment(line):
    q = False
    for i, ch in enumerate(line):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return line[:i]
    return line


def split_colon(line):
    parts, q, depth, cur = [], False, 0, ""
    for i, ch in enumerate(line):
        if ch == '"':
            q = not q
        if not q and ch == "(":
            depth += 1
        if not q and ch == ")":
            depth -= 1
        if ch == ":" and not q and depth == 0 and line[i + 1:i + 2] != "=":
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    parts.append(cur)
    return [p.strip() for p in parts if p.strip()]


def pyname(n):
    n = n.lower()
    return n + "_" if n in PYKW else n


class Module:
    def __init__(self, name, text):
        self.name = name
        self.vars, self.consts, self.procs = {}, {}, {}   # nom -> (public?)
        self.body = []                                    # (kind, header, lines)
        self.decls = []
        lines = text.replace("\r\n", "\n").split("\n")
        # davom etuvchi qatorlarni birlashtirish
        joined, buf = [], ""
        for ln in lines:
            ln = strip_comment(ln).rstrip()
            if ln.endswith(" _"):
                buf += ln[:-2] + " "
                continue
            joined.append(buf + ln)
            buf = ""
        i = 0
        while i < len(joined):
            ln = joined[i].strip()
            low = ln.lower()
            m = re.match(r"(public |private )?(sub|function) (\w+)\s*\((.*)\)(\s+as\s+\w+)?$", ln, re.I)
            if m:
                kind, pname = m.group(2).lower(), m.group(3)
                body = []
                i += 1
                while not re.match(r"end (sub|function)$", joined[i].strip(), re.I):
                    body.append(joined[i])
                    i += 1
                self.procs[pname.lower()] = (m.group(1) or "public ").strip().lower() == "public"
                self.body.append((kind, pname, m.group(4), body))
            elif low.startswith(("attribute", "option", "version", "begin", "end", "multiuse")) or not ln:
                pass
            else:
                self.decls.append(ln)
                m = re.match(r"(public |private |dim )?const (\w+)", ln, re.I)
                if m:
                    self.consts[m.group(2).lower()] = (m.group(1) or "").strip().lower() == "public"
                else:
                    m = re.match(r"(public|private|dim)\s+(.*)$", ln, re.I)
                    for part in split_args(m.group(2)):
                        nm = re.match(r"(\w+)", part.strip()).group(1)
                        self.vars[nm.lower()] = m.group(1).lower() == "public"
            i += 1


def split_args(s):
    out, depth, q, cur = [], 0, False, ""
    for ch in s:
        if ch == '"':
            q = not q
        if not q and ch == "(":
            depth += 1
        if not q and ch == ")":
            depth -= 1
        if ch == "," and depth == 0 and not q:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur)
    return out


def init_for(decl):
    """'x As Double' / 'xs() As Double' / 'xs(0 To 4) As Double' -> (nom, python boshlang'ich qiymati)"""
    m = re.match(r"(\w+)\s*(\((.*?)\))?\s*(as\s+(new\s+)?(\w+))?", decl.strip(), re.I)
    nm, dims, typ = m.group(1), m.group(3), (m.group(6) or "variant").lower()
    if m.group(2) is not None:
        if dims and dims.strip():
            lo, hi = re.split(r"\s+to\s+", dims, flags=re.I) if re.search(r"\sto\s", dims, re.I) else ("0", dims)
            return nm, ("arr", lo, hi)
        return nm, ("arr", None, None)
    return nm, {"double": "0.0", "single": "0.0", "long": "0", "integer": "0", "string": 'VStr("")',
                "boolean": "False"}.get(typ, "None")


class Translator:
    def __init__(self, modules):
        self.mods = {m.name.lower(): m for m in modules}

    def resolve(self, name, mod, local):
        n = name.lower()
        if n in local:
            return pyname(n)
        if n in OPS:
            return OPS[n]
        m = self.mods[mod]
        if n in m.vars or n in m.consts:
            return "G_%s.%s" % (mod, pyname(n))
        if n in m.procs:
            return "%s__%s" % (mod, pyname(n))
        for k, o in self.mods.items():
            if n in o.vars and o.vars[n] or n in o.consts and o.consts[n]:
                return "G_%s.%s" % (k, pyname(n))
            if n in o.procs and o.procs[n]:
                return "%s__%s" % (k, pyname(n))
        if n in RUNTIME:
            return n
        raise NameError("%s: e'lon qilinmagan nom '%s'" % (mod, name))

    def is_proc(self, name, mod, local):
        r = self.resolve(name, mod, local) if name.lower() not in local else ""
        return "__" in r

    def expr(self, s, mod, local, fnname=None):
        toks = [t for t in tokens(s) if not t.isspace()]
        out = []
        for i, t in enumerate(toks):
            prev = toks[i - 1] if i else ""
            nxt = toks[i + 1] if i + 1 < len(toks) else ""
            if t.startswith('"'):
                out.append('VStr(%s)' % ('"' + t[1:-1].replace('""', '\\"') + '"'))
            elif re.match(r"[A-Za-z_]", t):
                if prev == ".":
                    out.append(t.lower())
                elif fnname and t.lower() == fnname.lower() and nxt != "(":
                    out.append("_ret")
                else:
                    r = self.resolve(t, mod, local)
                    if "__" in r and nxt != "(":
                        r += "()"
                    out.append(r)
            elif t == "=":
                out.append("==")
            elif t == "<>":
                out.append("!=")
            elif t == "&":
                out.append("|")
            elif t == "\\":
                out.append("//")
            elif t == "^":
                out.append("**")
            elif t == ":=":
                raise SyntaxError("nomli argumentlar qo'llanmaydi: " + s)
            else:
                out.append(t)
        return " ".join(out).replace(" . ", ".").replace("( ", "(").replace(" )", ")").replace(" ,", ",")

    def proc(self, kind, name, params, body, mod):
        local, pyparams, decl_seen = set(), [], set()
        for p in split_args(params):
            p = p.strip()
            pm = re.match(r"(optional\s+)?(byval\s+|byref\s+)?(paramarray\s+)?(\w+)(\(\))?(\s+as\s+\w+)?(\s*=\s*(.+))?$", p, re.I)
            nm = pm.group(4)
            if nm.lower() in decl_seen:
                raise SyntaxError("%s.%s: takroriy nom %s" % (mod, name, nm))
            decl_seen.add(nm.lower())
            local.add(nm.lower())
            if pm.group(3):
                pyparams.append("*" + pyname(nm))
            elif pm.group(1):
                pyparams.append("%s=%s" % (pyname(nm), self.expr(pm.group(8), mod, set())))
            else:
                pyparams.append(pyname(nm))
        # mahalliy e'lonlarni oldindan yig'ish (VBA: protsedura darajasida)
        for ln in body:
            for st in split_colon(ln):
                m = re.match(r"(dim|const)\s+(.*)$", st, re.I)
                if m:
                    for part in split_args(m.group(2)):
                        nm = re.match(r"\s*(\w+)", part).group(1)
                        if nm.lower() in decl_seen:
                            raise SyntaxError("%s.%s: takroriy e'lon %s" % (mod, name, nm))
                        decl_seen.add(nm.lower())
                        local.add(nm.lower())
        out = ["def %s__%s(%s):" % (mod, pyname(name), ", ".join(pyparams))]
        ind = 1
        for pp in pyparams:
            if pp.startswith("*"):
                out.append("    %s = PArr(%s)" % (pp[1:], pp[1:]))
        if kind == "function":
            out.append("    _ret = 0")
        sel = []

        def emit(code):
            out.append("    " * ind + code)

        def stmt(st):
            nonlocal ind
            low = st.lower()
            E = lambda s: self.expr(s, mod, local, name if kind == "function" else None)
            m = re.match(r"if (.+) then(.*)$", st, re.I)
            if m and m.group(2).strip():                       # bir qatorli If
                cond, rest = m.group(1), m.group(2).strip()
                parts = re.split(r"\belse\b", rest, maxsplit=1, flags=re.I)
                emit("if %s:" % E(cond))
                ind += 1
                for s2 in split_colon(parts[0]):
                    stmt(s2)
                ind -= 1
                if len(parts) > 1:
                    emit("else:")
                    ind += 1
                    for s2 in split_colon(parts[1]):
                        stmt(s2)
                    ind -= 1
                return
            if m:
                emit("if %s:" % E(m.group(1)))
                ind += 1
                return
            m = re.match(r"elseif (.+) then$", st, re.I)
            if m:
                ind -= 1
                emit("elif %s:" % E(m.group(1)))
                ind += 1
                return
            if low == "else":
                ind -= 1
                emit("else:")
                ind += 1
                return
            if low in ("end if", "next") or re.match(r"next \w+$", low) or low == "loop":
                ind -= 1
                return
            m = re.match(r"for (\w+) = (.+?) to (.+?)( step (.+))?$", st, re.I)
            if m:
                emit("for %s in _rng(%s, %s%s):" % (pyname(m.group(1)), E(m.group(2)), E(m.group(3)),
                                                   ", " + E(m.group(5)) if m.group(5) else ""))
                ind += 1
                return
            m = re.match(r"for each (\w+) in (.+)$", st, re.I)
            if m:
                emit("for %s in %s:" % (pyname(m.group(1)), E(m.group(2))))
                ind += 1
                return
            m = re.match(r"do while (.+)$", st, re.I)
            if m:
                emit("while %s:" % E(m.group(1)))
                ind += 1
                return
            m = re.match(r"select case (.+)$", st, re.I)
            if m:
                v = "_sel%d" % len(sel)
                emit("%s = %s" % (v, E(m.group(1))))
                sel.append([v, True])
                return
            m = re.match(r"case (.+)$", st, re.I)
            if m:
                v, first = sel[-1]
                if not first:
                    ind -= 1
                if m.group(1).lower() == "else":
                    emit("else:" if not first else "if True:")
                else:
                    conds = []
                    for c in split_args(m.group(1)):
                        r = re.split(r"\s+to\s+", c.strip(), flags=re.I)
                        conds.append("(%s <= %s <= %s)" % (E(r[0]), v, E(r[1])) if len(r) == 2 else "%s == %s" % (v, E(r[0])))
                    emit("%s %s:" % ("if" if first else "elif", " or ".join(conds)))
                sel[-1][1] = False
                ind += 1
                return
            if low == "end select":
                if not sel.pop()[1]:
                    ind -= 1
                return
            if low in ("exit sub", "exit function"):
                emit("return _ret" if kind == "function" else "return")
                return
            if low in ("exit for", "exit do"):
                emit("break")
                return
            m = re.match(r"(dim|const)\s+(.*)$", st, re.I)
            if m:
                if m.group(1).lower() == "const":
                    cm = re.match(r"(\w+)(\s+as\s+\w+)?\s*=\s*(.+)$", m.group(2), re.I)
                    emit("%s = %s" % (pyname(cm.group(1)), E(cm.group(3))))
                    return
                for part in split_args(m.group(2)):
                    nm, init = init_for(part)
                    if isinstance(init, tuple):
                        emit("%s = VArr(%s, %s)" % (pyname(nm), E(init[1]), E(init[2])) if init[1] else "%s = VArr()" % pyname(nm))
                    else:
                        emit("%s = %s" % (pyname(nm), init))
                return
            m = re.match(r"redim (preserve )?(\w+)\((.+)\)$", st, re.I)
            if m:
                dims = m.group(3)
                lo, hi = re.split(r"\s+to\s+", dims, flags=re.I) if re.search(r"\sto\s", dims, re.I) else ("0", dims)
                emit("%s.redim(%s, %s, %s)" % (E(m.group(2)), E(lo), E(hi), bool(m.group(1))))
                return
            if low.startswith("set "):
                st = st[4:]
            # o'zlashtirish:  nom[(...)][.a.b] = ifoda
            toks = tokens(st)
            depth, eq = 0, -1
            for i, t in enumerate(toks):
                depth += t == "("
                depth -= t == ")"
                if t == "=" and depth == 0:
                    eq = i
                    break
            if eq > 0:
                lhs, rhs = "".join(toks[:eq]).strip(), "".join(toks[eq + 1:])
                lm = re.match(r"(\w+)\s*\((.+)\)$", lhs)
                if kind == "function" and lhs.lower() == name.lower():
                    emit("_ret = %s" % E(rhs))
                elif lm:
                    emit("%s[%s] = %s" % (E(lm.group(1)).replace("()", ""), E(lm.group(2)), E(rhs)))
                else:
                    emit("%s = %s" % (E(lhs).replace("()", ""), E(rhs)))
                return
            # protsedura chaqiruvi:  Nom arg, arg   yoki  obj.Metod arg
            cm = re.match(r"([\w.]+)\s*(.*)$", st)
            if cm:
                head, args = cm.group(1), cm.group(2)
                emit("%s(%s)" % (E(head).replace("()", ""), ", ".join(E(a) for a in split_args(args)) if args else ""))
                return
            raise SyntaxError("tushunilmadi: " + st)

        for ln in body:
            if re.match(r"\s*[A-Za-z_]\w*:(\s|$)", ln) and not re.match(r"\s*(case|else)\b", ln, re.I):
                raise SyntaxError("%s.%s: VBA buni label deb tushunadi (qator boshida 'Nom:'): %s" % (mod, name, ln.strip()))
            parts = split_colon(ln.strip())
            for i, st in enumerate(parts):
                m = re.match(r"if .+? then\s*\S", st, re.I)
                if m:                                  # bir qatorli If qatorning oxirigacha davom etadi
                    stmt(": ".join(parts[i:]))
                    break
                stmt(st)
        if kind == "function":
            out.append("    return _ret")
        if len(out) == 1:
            out.append("    pass")
        return "\n".join(out)

    def run(self):
        code = []
        for k, m in self.mods.items():
            code.append("class G_%s: pass" % k)
        for k, m in self.mods.items():
            for d in m.decls:
                mm = re.match(r"(public |private |dim )?const (\w+)(\s+as\s+\w+)?\s*=\s*(.+)$", d, re.I)
                if mm:
                    code.append("G_%s.%s = %s" % (k, pyname(mm.group(2)), self.expr(mm.group(4), k, set())))
                    continue
                mm = re.match(r"(public|private|dim)\s+(.*)$", d, re.I)
                for part in split_args(mm.group(2)):
                    nm, init = init_for(part)
                    code.append("G_%s.%s = %s" % (k, pyname(nm), "VArr()" if isinstance(init, tuple) else init))
            for kind, name, params, body in m.body:
                code.append(self.proc(kind, name, params, body, k))
        return "\n\n".join(code)


def load(paths, skip=()):
    """VBA fayllarni o'girib, Python nom fazosini qaytaradi"""
    mods = []
    for p in paths:
        text = open(p, encoding="latin-1").read()
        name = re.search(r'Attribute VB_Name = "(\w+)"', text).group(1)
        m = Module(name, text)
        m.body = [b for b in m.body if b[1] not in skip]
        mods.append(m)
    src = Translator(mods).run()
    ns = dict(RUNTIME)
    exec(compile(src, "<vba>", "exec"), ns)
    return ns, src
