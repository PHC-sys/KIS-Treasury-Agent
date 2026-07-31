# -*- coding: utf-8 -*-
"""
db/store.py — ledger 접근 계층 (멱등 upsert · 최종일 조회 · wide export)

멱등성의 심장: upsert는 (date, field) 충돌 시 덮어쓰기.
같은 데이터를 몇 번 넣어도 결과 동일.
"""
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import config


def connect():
    config.DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(config.DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def init_db(conn):
    schema = (Path(__file__).parent / "schema.sql").read_text(encoding="utf-8")
    conn.executescript(schema)
    conn.commit()


def now_iso():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def upsert(conn, rows):
    """rows: iterable of (date, field, value, source, as_of, fetched_at).
    (date, field) 충돌 시 덮어쓴다 → 멱등."""
    conn.executemany(
        """
        INSERT INTO ledger (date, field, value, source, as_of, fetched_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(date, field) DO UPDATE SET
            value=excluded.value,
            source=excluded.source,
            as_of=excluded.as_of,
            fetched_at=excluded.fetched_at
        """,
        list(rows),
    )
    conn.commit()


def quarantine(conn, rows):
    """rows: iterable of (date, field, value, source, reason, fetched_at)."""
    conn.executemany(
        """
        INSERT OR REPLACE INTO quarantine
            (date, field, value, source, reason, fetched_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        list(rows),
    )
    conn.commit()


def last_date_for_source(conn, source):
    """해당 소스가 적재한 마지막 date. 없으면 None (최초 실행 판단용)."""
    fields = config.SOURCE_FIELDS.get(source, [])
    if not fields:
        return None
    qs = ",".join("?" * len(fields))
    row = conn.execute(
        f"SELECT MAX(date) FROM ledger WHERE field IN ({qs})", fields
    ).fetchone()
    return row[0] if row and row[0] else None


def all_dates(conn):
    """ledger에 존재하는 모든 date (거래일) 오름차순."""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT date FROM ledger ORDER BY date").fetchall()]


def field_series(conn, field):
    """{date: value} — 파생계산(누적합)·동결감지용."""
    return {r[0]: r[1] for r in conn.execute(
        "SELECT date, value FROM ledger WHERE field=? ORDER BY date",
        (field,)).fetchall()}


def wide_table(conn):
    """ledger → wide dict: {date: {field: value}}. 결측은 키 없음(빈칸)."""
    table = {}
    for date, field, value in conn.execute(
            "SELECT date, field, value FROM ledger"):
        table.setdefault(date, {})[field] = value
    return table
