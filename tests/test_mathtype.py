from eduq import mathtype as mt


def fake_mtef(*codes: int) -> bytes:
    """MTEF giả: phần đầu + 'MT Extra' + các record CHAR (typeface 0x88)."""
    body = b"".join(bytes([2, 0, 0x88, c & 0xFF, c >> 8]) for c in codes)
    return b"\x05\x01\x00\x07\x08DSMT7\x00\x01" + b"MT Extra\x00" + b"\x12\x00\x08" + body + b"\x00\x00"


def test_strip_outer_braces():
    assert mt.normalize_tex(r"{\frac{1}{2}}")[0] == r"\frac{1}{2}"
    assert mt.normalize_tex(r"{a}{b}")[0] == r"{a}{b}"  # không phải một cặp bao toàn bộ
    assert mt.normalize_tex("{{x}}")[0] == "x"


def test_degree_fix():
    s, notes = mt.normalize_tex(r"{\widehat{xOy} = 50\circ}")
    assert s == r"\widehat{xOy} = 50^{\circ}" and "fix:degree" in notes
    assert mt.normalize_tex(r"{50^{\circ}}")[0] == r"50^{\circ}"  # đã đúng thì giữ
    assert mt.normalize_tex(r"{50^\circ}")[0] == r"50^\circ"


def test_prime_fix():
    assert mt.normalize_tex(r"{xx\prime}")[0] == "xx'"


def test_scan_chars_reads_records():
    assert mt.scan_chars(fake_mtef(0x35, 0x30, 0xB0)) == [0x35, 0x30, 0xB0]


def test_crosscheck_ok():
    assert mt.crosscheck(r"50^{\circ}", fake_mtef(0x35, 0x30, 0xB0)) == []


def test_crosscheck_detects_missing_degree_symbol():
    # TeX gốc `50\circ` đã được sửa; nhưng nếu LaTeX mất hẳn ký hiệu thì phải báo
    issues = mt.crosscheck("50", fake_mtef(0x35, 0x30, 0xB0))
    assert any(i.startswith("symbol_missing:U+00B0") for i in issues)


def test_crosscheck_detects_wrong_digits():
    issues = mt.crosscheck("51", fake_mtef(0x35, 0x30))
    assert any(i.startswith("alnum_mismatch") for i in issues)


def test_crosscheck_flags_unknown_symbol():
    issues = mt.crosscheck("x", fake_mtef(0x78, 0x2A01))
    assert "unknown_symbol:U+2A01" in issues


def test_crosscheck_blackboard_letters():
    assert mt.crosscheck(r"\mathbb{Q}", fake_mtef(0x211A)) == []


def test_crosscheck_without_mtef_is_flagged():
    assert mt.crosscheck("x", None) == ["no_mtef"]


def test_read_mtef_rejects_garbage():
    assert mt.read_mtef(b"not an ole file") is None
