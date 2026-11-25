"""Kho tài sản (ảnh, OLE, preview) lưu theo SHA-256: tự khử trùng lặp, đường dẫn ổn định."""
from __future__ import annotations

from pathlib import Path


class AssetStore:
    def __init__(self, root: str | Path):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)

    def path(self, sha256: str) -> Path:
        return self.root / sha256[:2] / sha256[2:]

    def put(self, sha256: str, data: bytes) -> Path:
        p = self.path(sha256)
        if not p.exists():
            p.parent.mkdir(parents=True, exist_ok=True)
            tmp = p.with_suffix(".tmp")
            tmp.write_bytes(data)
            tmp.replace(p)
        return p

    def read(self, sha256: str) -> bytes:
        return self.path(sha256).read_bytes()
