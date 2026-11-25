"""Mô hình dữ liệu (SQLAlchemy 2.0). Chạy SQLite khi phát triển, đổi URL sang PostgreSQL khi triển khai."""
from __future__ import annotations

from datetime import datetime, timezone

from sqlalchemy import JSON, DateTime, ForeignKey, Integer, String, Text, UniqueConstraint, create_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship, sessionmaker


def _now() -> datetime:
    return datetime.now(timezone.utc)


class Base(DeclarativeBase):
    pass


class Topic(Base):
    __tablename__ = "topics"
    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(200), unique=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=_now)


class SourceFile(Base):
    __tablename__ = "source_files"
    __table_args__ = (UniqueConstraint("topic_id", "sha256"),)
    id: Mapped[int] = mapped_column(primary_key=True)
    topic_id: Mapped[int] = mapped_column(ForeignKey("topics.id"))
    filename: Mapped[str] = mapped_column(String(500))
    sha256: Mapped[str] = mapped_column(String(64))
    imported_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=_now)
    report: Mapped[dict] = mapped_column(JSON, default=dict)
    topic: Mapped[Topic] = relationship()


class Asset(Base):
    __tablename__ = "assets"
    sha256: Mapped[str] = mapped_column(String(64), primary_key=True)
    kind: Mapped[str] = mapped_column(String(20))  # image | ole | preview
    mime: Mapped[str] = mapped_column(String(100))
    size: Mapped[int] = mapped_column(Integer)
    name: Mapped[str] = mapped_column(String(300))


class Question(Base):
    """Một bài tập trong kho. Mã `C0000001` sinh từ id tự tăng, không bao giờ tái sử dụng."""

    __tablename__ = "questions"
    __table_args__ = {"sqlite_autoincrement": True}
    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    code: Mapped[str | None] = mapped_column(String(16), unique=True)
    topic_id: Mapped[int] = mapped_column(ForeignKey("topics.id"), index=True)
    source_file_id: Mapped[int] = mapped_column(ForeignKey("source_files.id"))
    exam: Mapped[str | None] = mapped_column(String(200))
    section: Mapped[str | None] = mapped_column(String(300))
    number: Mapped[int] = mapped_column(Integer)  # thứ tự trong đề
    qtype: Mapped[str] = mapped_column(String(20))
    points: Mapped[float | None]
    content: Mapped[dict] = mapped_column(JSON)  # {"stem": [...], "options": [...] | None}
    answer: Mapped[dict | None] = mapped_column(JSON, default=None)  # nhập sau (web hoặc file đáp án)
    raw_ooxml: Mapped[str] = mapped_column(Text)  # OOXML thô để xuất lại y nguyên
    raw_rels: Mapped[dict] = mapped_column(JSON)  # rId → {type, target, asset}
    status: Mapped[str] = mapped_column(String(20), default="ok", index=True)  # ok | needs_review
    issues: Mapped[list] = mapped_column(JSON, default=list)
    content_hash: Mapped[str] = mapped_column(String(64), index=True)
    duplicate_of: Mapped[str | None] = mapped_column(String(16))
    topic: Mapped[Topic] = relationship()


def make_code(n: int) -> str:
    return f"C{n:07d}"


def open_db(url: str):
    engine = create_engine(url)
    Base.metadata.create_all(engine)
    return sessionmaker(engine, expire_on_commit=False)
