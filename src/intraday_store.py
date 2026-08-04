# -*- coding: utf-8 -*-
"""
intraday_store.py — 5분봉(선물·현물) 저장 계층.

일별 ledger(store.py)와 별도 파일(db/intraday.sqlite)·별도 형태.
일별은 EAV(long) 스칼라, 인트라데이는 wide 1행/봉 (OHLCV+OI).

규율 (인수인계 §8a):
- **증분 전용**: 종목별 마지막 ts 이후만 insert. 전체 재수집 금지.
- 멱등: PRIMARY KEY(symbol, ts) 충돌 시 덮어씀 → 몇 번 돌려도 동일.
- 결측=NULL, 0은 진짜 0. 선물=가격 / 현물=수익률(%).
"""
import sqlite3
from datetime import datetime, timezone

import config

SCHEMA = """
CREATE TABLE IF NOT EXISTS bars (
    symbol     TEXT NOT NULL,   -- C65/C67(선물) | KR1035G000xx(현물 지표연속)
    ts         TEXT NOT NULL,   -- 'YYYY-MM-DD HH:MM:SS' (일자+시간 결합, 봉 시각)
    open       REAL,
    high       REAL,
    low        REAL,
    close      REAL,            -- 선물=가격 / 현물=수익률(%)
    volume     REAL,            -- 선물 거래량 / 현물 체결거래량(액면)
    oi         REAL,            -- 선물 미결제약정수량 (현물 NULL)
    fetched_at TEXT,
    PRIMARY KEY (symbol, ts)
);
"""
# ※ 체결거래대금(amount)은 지표연속에 IMDH가 미산출(항상 0) → 컬럼 제거(2026-08-04).


def connect():
    config.INTRADAY_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(config.INTRADAY_DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def init_db(conn):
    conn.executescript(SCHEMA)
    conn.commit()


def now_iso():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def upsert(conn, rows):
    """rows: iterable of (symbol, ts, open, high, low, close, volume, oi, fetched_at).
    (symbol, ts) 충돌 시 덮어씀 → 멱등. 재실행/부분겹침 백필에도 안전."""
    conn.executemany(
        """
        INSERT INTO bars (symbol, ts, open, high, low, close, volume, oi, fetched_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(symbol, ts) DO UPDATE SET
            open=excluded.open, high=excluded.high, low=excluded.low, close=excluded.close,
            volume=excluded.volume, oi=excluded.oi,
            fetched_at=excluded.fetched_at
        """,
        list(rows),
    )
    conn.commit()


def last_ts(conn, symbol):
    """해당 종목이 저장한 마지막 봉 ts (증분 창의 시작 판단용). 없으면 None → 전체 백필."""
    row = conn.execute(
        "SELECT MAX(ts) FROM bars WHERE symbol=?", (symbol,)
    ).fetchone()
    return row[0] if row and row[0] else None


def first_ts(conn, symbol):
    """해당 종목의 가장 오래된 봉 ts (이력 깊이 확인용)."""
    row = conn.execute(
        "SELECT MIN(ts) FROM bars WHERE symbol=?", (symbol,)
    ).fetchone()
    return row[0] if row and row[0] else None


def count(conn, symbol=None):
    if symbol is None:
        return conn.execute("SELECT COUNT(*) FROM bars").fetchone()[0]
    return conn.execute(
        "SELECT COUNT(*) FROM bars WHERE symbol=?", (symbol,)
    ).fetchone()[0]


def coverage(conn):
    """종목별 {symbol: (first_ts, last_ts, n)} — 백필/증분 진행 점검용."""
    out = {}
    for sym, lo, hi, n in conn.execute(
        "SELECT symbol, MIN(ts), MAX(ts), COUNT(*) FROM bars GROUP BY symbol"
    ):
        out[sym] = (lo, hi, n)
    return out
