"""Công thức MathType (OLE) → LaTeX, có kiểm chéo để không âm thầm đưa ra kết quả sai.

Hai đường chuyển (kết quả spike trên file thật: 385 + 12 = 397/397 công thức):
  1a. MathType 7 nhúng sẵn mã TeX trong stream `Equation Native` (record "TeX Input Language").
      Nhanh, giữ nguyên cấu trúc, nhưng MathType xuất một số ký hiệu thiếu chính xác
      (dấu độ ° bị ghi `\\circ`) → áp bảng sửa + kiểm chéo.
  1b. Không có TeX nhúng → MTEF → MathML (gem Ruby `mathtype_to_mathml`) → LaTeX (pandoc).
Kiểm chéo: chữ/số và ký hiệu đặc biệt trong LaTeX phải khớp với ký tự thật (mtcode) trong MTEF.
Không khớp / ký hiệu lạ → status = needs_review (vẫn giữ ảnh preview để hiển thị).
"""
from __future__ import annotations

import io
import json
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import olefile

TOOLS_DIR = Path(__file__).resolve().parents[2] / "tools"

HEADER_FALLBACK_LEN = 28
_TEX_RE = re.compile(rb"TeX Input Language\x00(.*?)\x00", re.S)

# ---------------------------------------------------------------- đọc OLE / MTEF


def read_mtef(blob: bytes) -> bytes | None:
    """Lấy phần MTEF (bỏ header 28 byte) từ OLE object `oleObjectN.bin`."""
    try:
        ole = olefile.OleFileIO(io.BytesIO(blob))
        try:
            if not ole.exists("Equation Native"):
                return None
            data = ole.openstream("Equation Native").read()
        finally:
            ole.close()
    except Exception:
        return None
    if len(data) < 4:
        return None
    hdr = int.from_bytes(data[:2], "little")
    if not 8 <= hdr <= 64:
        hdr = HEADER_FALLBACK_LEN
    return data[hdr:]


def extract_embedded_tex(mtef: bytes) -> str | None:
    m = _TEX_RE.search(mtef)
    if not m:
        return None
    return m.group(1).decode("utf-8", "replace")


def scan_chars(mtef: bytes) -> list[int]:
    """Quét thô các record CHAR của MTEF5 và trả về danh sách mã ký tự (mtcode).

    Chỉ dùng để kiểm chéo (không phải parser đầy đủ): sai sót chỉ gây cảnh báo, không gây sai im lặng.
    """
    start = mtef.find(b"MT Extra")
    b = mtef[start + 9 :] if start >= 0 else mtef
    out, i = [], 0
    while i < len(b) - 4:
        if b[i] == 2 and b[i + 1] < 0x20 and b[i + 2] >= 0x7F:
            attr = b[i + 1]
            code = b[i + 3] | (b[i + 4] << 8)
            j = i + 5
            if attr & 0x08:  # nudge
                j += 2
            if attr & 0x04:  # ký tự 8-bit kèm theo
                if b[i + 2] == 0x7F and j < len(b):
                    code = b[j]  # font người dùng (vd. Euclid Math Two): mã thật nằm ở byte 8-bit
                j += 1
            if attr & 0x10:  # ký tự 16-bit kèm theo
                j += 2
            out.append(code)
            i = j
        else:
            i += 1
    return out


# ---------------------------------------------------------------- chuẩn hóa LaTeX


def _strip_outer_braces(s: str) -> str:
    s = s.strip()
    while s.startswith("{") and s.endswith("}"):
        depth = 0
        for i, ch in enumerate(s):
            if ch == "{" and (i == 0 or s[i - 1] != "\\"):
                depth += 1
            elif ch == "}" and (i == 0 or s[i - 1] != "\\"):
                depth -= 1
                if depth == 0 and i != len(s) - 1:
                    return s  # dấu { đầu đóng trước cuối: không phải ngoặc bao toàn bộ
        s = s[1:-1].strip()
    return s


def normalize_tex(tex: str) -> tuple[str, list[str]]:
    """Bỏ ngoặc bao ngoài và sửa các ký hiệu MathType xuất thiếu chính xác."""
    notes: list[str] = []
    s = _strip_outer_braces(tex)
    # Dấu độ: MathType ghi ° (U+00B0) thành \circ (toán tử hợp thành, vòng to) → ^{\circ}.
    s2 = re.sub(r"(?<!\^)(?<!\^\{)\\circ(?![A-Za-z])", r"^{\\circ}", s)
    if s2 != s:
        notes.append("fix:degree")
        s = s2
    # Dấu phẩy-nháy (x') : \prime đứng riêng → '
    s2 = re.sub(r"\\prime(?![A-Za-z])", "'", s)
    if s2 != s:
        notes.append("fix:prime")
        s = s2
    return s, notes


# ---------------------------------------------------------------- kiểm chéo

# mtcode → regex phải xuất hiện trong LaTeX (đã chuẩn hóa) nếu MTEF có ký tự đó.
_GREEK = {
    0x3B1: "alpha", 0x3B2: "beta", 0x3B3: "gamma", 0x3B4: "delta", 0x3B5: "(?:varepsilon|epsilon)",
    0x3B8: "theta", 0x3BB: "lambda", 0x3BC: "mu", 0x3C0: "pi", 0x3C1: "rho", 0x3C3: "sigma",
    0x3C6: "(?:varphi|phi)", 0x3C9: "omega", 0x394: "Delta", 0x3A3: "Sigma", 0x3A9: "Omega",
}
SYMBOL_EXPECT: dict[int, str] = {
    0x2212: r"-", 0x2C: r",", 0x3D: r"=", 0x2B: r"\+", 0x3A: r":", 0x2E: r"\.", 0x3B: r";",
    0x28: r"\(", 0x29: r"\)", 0x5B: r"\[", 0x5D: r"\]", 0x2F: r"/", 0x21: r"!", 0x3C: r"<",
    0x3E: r">", 0x7C: r"\|", 0x25: r"%", 0x20: r"", 0xA0: r"",
    0xB0: r"\\circ",
    0x22C5: r"\\cdot", 0x2225: r"\\parallel", 0x2032: r"'|\\prime", 0x22A5: r"\\(?:perp|bot)",
    0xB1: r"\\pm", 0x2261: r"\\equiv", 0x2248: r"\\approx", 0x302: r"\\(?:widehat|hat)",
    0xD7: r"\\times", 0xF7: r"\\div", 0x2264: r"\\(?:le|leq)", 0x2265: r"\\(?:ge|geq)",
    0x2260: r"\\(?:ne|neq)", 0x221E: r"\\infty", 0x2229: r"\\cap", 0x222A: r"\\cup",
    0x2208: r"\\in", 0x2209: r"\\notin", 0x2282: r"\\subset", 0x2286: r"\\subseteq",
    0x211A: r"\\mathbb\{Q\}", 0x2124: r"\\mathbb\{Z\}", 0x2115: r"\\mathbb\{N\}",
    0x211D: r"\\mathbb\{R\}", 0x2102: r"\\mathbb\{C\}", 0x1D540: r"\\mathbb\{I\}",
    0x2220: r"\\angle", 0x394: r"\\Delta", 0x25B3: r"\\triangle",
}
for _code, _name in _GREEK.items():
    SYMBOL_EXPECT.setdefault(_code, r"\\" + _name)

_BBB = re.compile(r"\\mathbb\{([A-Za-z])\}")
_BBB_CODE_LETTER = {0x211A: "Q", 0x2124: "Z", 0x2115: "N", 0x211D: "R", 0x2102: "C"}


def crosscheck(latex: str, mtef: bytes | None) -> list[str]:
    """Trả về danh sách cảnh báo (rỗng = khớp). So LaTeX với ký tự thật trong MTEF."""
    if mtef is None:
        return ["no_mtef"]
    codes = scan_chars(mtef)
    if not codes:
        return ["mtef_scan_empty"]
    problems: list[str] = []
    # 1) chữ/số theo thứ tự
    mt_seq = "".join(
        _BBB_CODE_LETTER.get(c) or (chr(c) if c < 128 and chr(c).isalnum() else "") for c in codes
    )
    tx = _BBB.sub(r"\1", latex)
    tx = re.sub(r"\\[A-Za-z]+", "", tx)
    tx_seq = "".join(ch for ch in tx if ch.isascii() and ch.isalnum())
    if mt_seq != tx_seq:
        problems.append(f"alnum_mismatch:mtef={mt_seq!r} latex={tx_seq!r}")
    # 2) ký hiệu đặc biệt phải có mặt
    for code in sorted({c for c in codes if not (c < 128 and chr(c).isalnum()) and c not in _BBB_CODE_LETTER}):
        pat = SYMBOL_EXPECT.get(code)
        if pat is None:
            problems.append(f"unknown_symbol:U+{code:04X}")
        elif pat and not re.search(pat, latex):
            problems.append(f"symbol_missing:U+{code:04X}")
    return problems


# ---------------------------------------------------------------- đường 1b: MTEF → MathML → LaTeX


def mathml_available() -> bool:
    if not (shutil.which("ruby") and shutil.which("pandoc")):
        return False
    r = subprocess.run(
        ["ruby", "-e", "require 'mathtype_to_mathml'"], capture_output=True, text=True, timeout=60
    )
    return r.returncode == 0


def blobs_to_mathml(blobs: dict[str, bytes]) -> dict[str, str | None]:
    """Chuyển hàng loạt OLE .bin → MathML qua gem Ruby (1 tiến trình cho cả lô)."""
    if not blobs:
        return {}
    with tempfile.TemporaryDirectory() as td:
        names = {}
        for i, (key, data) in enumerate(blobs.items()):
            fn = f"{i}.bin"
            (Path(td) / fn).write_bytes(data)
            names[fn] = key
        r = subprocess.run(
            ["ruby", str(TOOLS_DIR / "mtef_to_mathml.rb"), *[str(Path(td) / n) for n in names]],
            capture_output=True, text=True, timeout=600,
        )
        if r.returncode != 0:
            return {k: None for k in blobs}
        try:
            res = json.loads(r.stdout.strip().splitlines()[-1])
        except (ValueError, IndexError):
            return {k: None for k in blobs}
    return {key: res.get(fn) for fn, key in names.items()}


def mathml_to_latex(mathml: str) -> str | None:
    body = re.sub(r"^<\?xml[^>]*\?>\s*", "", mathml).replace('<math display="block">', "<math>")
    r = subprocess.run(
        ["pandoc", "-f", "html", "-t", "latex"],
        input=f"<html><body>{body}</body></html>", capture_output=True, text=True, timeout=60,
    )
    if r.returncode != 0:
        return None
    out = r.stdout.strip()
    m = re.fullmatch(r"\\\((.*)\\\)", out, re.S)
    if not m:
        return None
    s = m.group(1).strip()
    return re.sub(r"\\hat\{", r"\\widehat{", s)  # góc hình học: dấu mũ rộng, nhất quán với đường 1a


# ---------------------------------------------------------------- điều phối


@dataclass
class Formula:
    latex: str | None
    source: str  # "mathtype-tex" | "mathtype-mathml" | "none"
    status: str  # "ok" | "needs_review"
    notes: list[str] = field(default_factory=list)


def convert_many(blobs: dict[str, bytes], *, use_mathml: bool = True) -> dict[str, Formula]:
    """Chuyển một lô OLE object (khóa tùy ý, thường là sha256) → Formula."""
    result: dict[str, Formula] = {}
    mtefs: dict[str, bytes | None] = {k: read_mtef(b) for k, b in blobs.items()}
    pending: dict[str, bytes] = {}
    for key, mtef in mtefs.items():
        tex = extract_embedded_tex(mtef) if mtef else None
        if tex is None:
            pending[key] = blobs[key]
            continue
        latex, notes = normalize_tex(tex)
        issues = crosscheck(latex, mtef)
        result[key] = Formula(
            latex=latex, source="mathtype-tex",
            status="ok" if not issues else "needs_review", notes=notes + issues,
        )
    if pending:
        mml = blobs_to_mathml(pending) if use_mathml and mathml_available() else {}
        for key in pending:
            latex = mathml_to_latex(mml[key]) if mml.get(key) else None
            if latex is None:
                result[key] = Formula(None, "none", "needs_review", ["no_converter_result"])
                continue
            issues = crosscheck(latex, mtefs[key])
            result[key] = Formula(
                latex=latex, source="mathtype-mathml",
                status="ok" if not issues else "needs_review", notes=issues,
            )
    return result
