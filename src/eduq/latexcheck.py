"""Kiểm tra LaTeX có render được bằng KaTeX (chế độ nghiêm ngặt) không, gọi 1 lần cho cả lô."""
from __future__ import annotations

import json
import shutil
import subprocess

from .mathtype import TOOLS_DIR


def katex_available() -> bool:
    return bool(shutil.which("node")) and (TOOLS_DIR / "node_modules" / "katex").is_dir()


def validate(latexes: list[str]) -> list[str | None] | None:
    """Trả về danh sách lỗi theo thứ tự (None = hợp lệ). None nếu không có node/katex."""
    if not latexes:
        return []
    if not katex_available():
        return None
    r = subprocess.run(
        ["node", str(TOOLS_DIR / "katex_check.js")],
        input=json.dumps(latexes), capture_output=True, text=True, timeout=300,
    )
    if r.returncode != 0:
        return None
    return json.loads(r.stdout)
