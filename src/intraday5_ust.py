# -*- coding: utf-8 -*-
"""
intraday5_ust.py — 미국채 2년/10년 5분봉 수집기 (IR/US02Y·US10Y).

★기존 코드 무수정 원칙: intraday_pull.py의 COM 헬퍼만 import해서 재사용하고,
  UST 전용 로직(24h·세션필터 없음·OHLC만·별도 테이블 ust_bars)은 여기서만 구현.
  KTB 선물(C65/C67) 5분봉은 기존 intraday_pull.py가 그대로 담당(bars 테이블).

시간대: IMDH 분봉 일자/시간은 **한국시간**(단말 로컬). 일별 as-of와 정합 검증됨
  (한국 T일 06:00 봉 = 미국 전일 종가 = 일별 ust 값). → 시프트/변환 불필요, 그대로 저장.

데이터: 금리는 거래량/OI 미산출 → OHLC만. 무호가봉(close=0)은 드롭.

모드 (CLI):
  probe [SYM...]      가용 이력 깊이 확인 (저장 안 함)
  backfill [SYM...]   최초 이력 적재 (count=99999 = 가용 전이력, ~2025-07~)
  sync (기본) [SYM...] 이력 없으면 백필, 있으면 증분(last_ts 이후만). run.bat에서 호출.

★IMDH 분봉 gotcha (intraday_pull과 동일): B1/D1 날짜창 무시·count만 유효·99999 상한.
  UST는 24h(288봉/일)라 99999 ≈ 최근 ~347거래일까지만. 무거운 콜은 종목마다 프레시 Excel.
"""
import datetime as dt
import os
import sqlite3
import sys

import config
# 기존 COM/파싱 헬퍼 재사용 (intraday_pull.py 무수정)
from intraday_pull import (
    _open_app, _close_app, _wait_calc, _kill_headless_excel, _com_retry,
    _fnum, _fdate, _ftime,
)

# UST 종목 정의 (config INTRADAY_SYMBOLS와 분리 — 세션/컬럼 규칙이 달라 별도 관리)
UST_SYMBOLS = {
    "US02Y": {"kind": "IR", "label": "미국채 2년"},
    "US10Y": {"kind": "IR", "label": "미국채 10년"},
}
ITEMS = ["일자", "시간", "시가", "고가", "저가", "현재가"]   # ★분봉은 시/고/저/현재가 (MID_*는 0 반환)
UST_TABLE = "ust_bars"


# ── 저장 계층 (별도 테이블 ust_bars, 같은 intraday.sqlite) ────────────
_SCHEMA = f"""
CREATE TABLE IF NOT EXISTS {UST_TABLE} (
    symbol     TEXT NOT NULL,   -- US02Y | US10Y
    ts         TEXT NOT NULL,   -- 'YYYY-MM-DD HH:MM:SS' 한국시간 (일자+시간)
    open       REAL, high REAL, low REAL, close REAL,   -- 수익률(%), 거래량/OI 없음
    fetched_at TEXT,
    PRIMARY KEY (symbol, ts)
);
"""


def connect():
    config.INTRADAY_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(config.INTRADAY_DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def init_db(conn):
    conn.executescript(_SCHEMA)
    conn.commit()


def _now_iso():
    from datetime import timezone
    return dt.datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def upsert(conn, rows):
    conn.executemany(
        f"""INSERT INTO {UST_TABLE} (symbol, ts, open, high, low, close, fetched_at)
            VALUES (?,?,?,?,?,?,?)
            ON CONFLICT(symbol, ts) DO UPDATE SET
              open=excluded.open, high=excluded.high, low=excluded.low,
              close=excluded.close, fetched_at=excluded.fetched_at""",
        list(rows))
    conn.commit()


def last_ts(conn, symbol):
    r = conn.execute(f"SELECT MAX(ts) FROM {UST_TABLE} WHERE symbol=?", (symbol,)).fetchone()
    return r[0] if r and r[0] else None


def coverage(conn):
    out = {}
    for sym, lo, hi, n in conn.execute(
            f"SELECT symbol, MIN(ts), MAX(ts), COUNT(*) FROM {UST_TABLE} GROUP BY symbol"):
        out[sym] = (lo, hi, n)
    return out


# ── 파싱 (24h 전량, 무호가봉 드롭, 세션필터 없음) ─────────────────────
def _parse(rows, symbol):
    fetched = _now_iso()
    out = []
    for r in rows:
        if not r or r[0] is None:
            continue
        d = _fdate(r[0])
        tm = _ftime(r[1]) if len(r) > 1 else None
        if not d or not tm:
            continue
        o, h, lo, c = (_fnum(r[i]) if i < len(r) else None for i in (2, 3, 4, 5))
        if (c or 0) == 0:                 # 무호가봉(금리 미갱신) 드롭
            continue
        out.append((symbol, f"{d} {tm}", o, h, lo, c, fetched))
    out.sort(key=lambda x: x[1])
    return out


# ── 단일 조회 (count만 유효, B1/D1 무시) ─────────────────────────────
def _pull(app, symbol, spec, count):
    wb = _com_retry(lambda: app.Workbooks.Add())
    ws = wb.Worksheets(1)
    try:
        ncols = len(ITEMS)
        ws.Range("B1").Value = dt.datetime(2000, 1, 1)
        ws.Range("D1").Value = dt.datetime.now()
        ws.Range("F1").Value = count
        for j, it in enumerate(ITEMS):
            ws.Cells(3, j + 1).Value = it
        last_col = chr(ord("A") + ncols - 1)
        opt = (f"Per=분,Cycle={config.INTRADAY_CYCLE},sort=A,real=false,Bizday=0,"
               f"Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true")
        ws.Range("A2").Formula = (
            f'=IMDH("{spec["kind"]}","{symbol}",A3:{last_col}3,$B$1,$D$1,$F$1,"{opt}")')
        n = _wait_calc(app, ws, tag=symbol)
        if n <= 3:
            return []
        used = ws.UsedRange.Rows.Count
        if used < 4:
            return []
        rng = ws.Range(ws.Cells(4, 1), ws.Cells(used, ncols)).Value
        rows = rng if isinstance(rng, tuple) else ((rng,),)
        return _parse(rows, symbol)
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def _symbols(names):
    if not names:
        return list(UST_SYMBOLS.items())
    return [(n, UST_SYMBOLS[n]) for n in names]


def _floor_str():
    return dt.datetime.fromisoformat(config.INTRADAY_BACKFILL_START).strftime("%Y-%m-%d %H:%M:%S")


def _ingest(conn, sym, spec, count, floor, newer_than=None):
    """종목 1개: 프레시 앱으로 count봉 조회 → floor/증분 필터 → upsert. (n, first, last, raw_min)."""
    _kill_headless_excel()
    app, blank = _open_app()
    try:
        bars = _pull(app, sym, spec, count)
    finally:
        _close_app(app, blank)
    if not bars:
        return 0, None, None, None
    raw_min = bars[0][1]
    bars = [b for b in bars if b[1] >= floor]
    if newer_than is not None:
        bars = [b for b in bars if b[1] > newer_than]
    if bars:
        upsert(conn, bars)
    return len(bars), (bars[0][1] if bars else None), (bars[-1][1] if bars else None), raw_min


def probe(names=None):
    try:
        for sym, spec in _symbols(names):
            _kill_headless_excel()
            app, blank = _open_app()
            try:
                bars = _pull(app, sym, spec, config.INTRADAY_BACKFILL_COUNT)
            finally:
                _close_app(app, blank)
            if bars:
                days = len({b[1][:10] for b in bars})
                print(f"  {sym:8} {spec['label']:10} {bars[0][1]} ~ {bars[-1][1]} | {len(bars)}봉/{days}일")
            else:
                print(f"  {sym:8} {spec['label']:10} 데이터 없음")
    finally:
        _kill_headless_excel()


def backfill(names=None):
    conn = connect()
    init_db(conn)
    floor = _floor_str()
    try:
        for sym, spec in _symbols(names):
            print(f"  ▶ backfill {sym} ({spec['label']})")
            n, f, l, _ = _ingest(conn, sym, spec, config.INTRADAY_BACKFILL_COUNT, floor)
            print(f"    = {sym}: {n}봉  {f} → {l}" if n else f"    = {sym}: 데이터 없음")
    finally:
        _kill_headless_excel()
    print("backfill 완료:", coverage(conn))
    conn.close()


def sync(names=None):
    """이력 없으면 백필(99999), 있으면 증분(last_ts 이후만). run.bat 진입점.
    ★UST 백필은 무거워(24h·99999) 프로세스당 1종목 권장 — 최초 1회는 CLI로 종목별 실행.
      일별 증분은 가벼워(INCR_COUNT) 한 프로세스 다종목 안전."""
    conn = connect()
    init_db(conn)
    floor = _floor_str()
    incr, full = config.INTRADAY_INCR_COUNT, config.INTRADAY_BACKFILL_COUNT
    try:
        for sym, spec in _symbols(names):
            last = last_ts(conn, sym)
            if last is None:
                print(f"  ▶ {sym} ({spec['label']}) 이력 없음 → 백필")
                n, f, l, _ = _ingest(conn, sym, spec, full, floor)
            else:
                print(f"  ▶ {sym} ({spec['label']}) 증분 (마지막 {last})")
                n, f, l, raw_min = _ingest(conn, sym, spec, incr, floor, newer_than=last)
                if raw_min is not None and raw_min > last:      # 공백이 커서 못 덮음 → 전체콜
                    print(f"    · 공백 큼(최소ts {raw_min} > {last}) → 전체콜 재수집")
                    n, f, l, _ = _ingest(conn, sym, spec, full, floor, newer_than=last)
            print(f"    = {sym}: +{n}봉  {f} → {l}" if n else f"    = {sym}: 신규 없음")
    finally:
        _kill_headless_excel()
    print("UST 동기화 완료:", coverage(conn))
    conn.close()


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "sync"
    names = sys.argv[2:] or None
    if mode in ("sync", "incremental"):
        sync(names)
    elif mode == "probe":
        probe(names)
    elif mode == "backfill":
        backfill(names)
    else:
        print(__doc__)
