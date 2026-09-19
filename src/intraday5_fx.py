# -*- coding: utf-8 -*-
"""
intraday5_fx.py — 달러원(USDKRW) 5분봉 수집기 (FX/USDSP_SMBCC_EXT).

★기존 코드 무수정 원칙: intraday_pull.py의 COM/파싱 헬퍼만 import해서 재사용.
  달러원 전용 로직(연장판 EXT·Close=중간값·OHLC만·별도 테이블 fx_bars)은 여기서만 구현.
  intraday5_ust.py(미국채)와 같은 패턴, 티커/옵션만 다름.

★티커 확정(2026-09-19 단말 실측):
  - 정규판 FX/USDSP_SMBCC 는 분봉 미지원("조회가 지원되지 않습니다").
  - **연장판 FX/USDSP_SMBCC_EXT + Close=중간값** 이라야 분봉이 나온다(24h/연장 포함).
  - 항목은 시가/고가/저가/현재가(중간값 기준). 거래량/OI 없음 → OHLC만.
저장 symbol = 'USDKRW'(테이블 표기용), IMDH 티커 = 'USDSP_SMBCC_EXT'.

시간대: 분봉 일자/시간은 한국시간(단말 로컬). 일별 usdkrw(정규장)와 별개 시리즈.

모드 (CLI):
  probe / backfill / sync (기본)  — intraday5_ust.py와 동일 의미.

★gotcha(intraday_pull과 동일): B1/D1 날짜창 무시·count만 유효·99999 상한.
  달러원은 단일 종목이라 백필 세그폴트 위험 낮음(종목 1개). infomax_data.xlsx 닫고 실행.
"""
import datetime as dt
import os
import sqlite3
import sys

import config
from intraday_pull import (
    _open_app, _close_app, _wait_calc, _kill_headless_excel, _com_retry,
    _fnum, _fdate, _ftime,
)

# 저장 symbol → IMDH (kind, ticker). 항목/옵션은 아래 공통.
FX_SYMBOLS = {
    "USDKRW": {"kind": "FX", "ticker": "USDSP_SMBCC_EXT", "label": "달러원 스팟(연장)"},
}
ITEMS = ["일자", "시간", "시가", "고가", "저가", "현재가"]   # Close=중간값 기준 OHLC
FX_TABLE = "fx_bars"


# ── 저장 계층 (별도 테이블 fx_bars, 같은 intraday.sqlite) ─────────────
_SCHEMA = f"""
CREATE TABLE IF NOT EXISTS {FX_TABLE} (
    symbol     TEXT NOT NULL,   -- USDKRW
    ts         TEXT NOT NULL,   -- 'YYYY-MM-DD HH:MM:SS' 한국시간
    open       REAL, high REAL, low REAL, close REAL,   -- 환율(중간값), 거래량/OI 없음
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
        f"""INSERT INTO {FX_TABLE} (symbol, ts, open, high, low, close, fetched_at)
            VALUES (?,?,?,?,?,?,?)
            ON CONFLICT(symbol, ts) DO UPDATE SET
              open=excluded.open, high=excluded.high, low=excluded.low,
              close=excluded.close, fetched_at=excluded.fetched_at""",
        list(rows))
    conn.commit()


def last_ts(conn, symbol):
    r = conn.execute(f"SELECT MAX(ts) FROM {FX_TABLE} WHERE symbol=?", (symbol,)).fetchone()
    return r[0] if r and r[0] else None


def coverage(conn):
    out = {}
    for sym, lo, hi, n in conn.execute(
            f"SELECT symbol, MIN(ts), MAX(ts), COUNT(*) FROM {FX_TABLE} GROUP BY symbol"):
        out[sym] = (lo, hi, n)
    return out


# ── 파싱 (24h/연장 전량, 무호가봉 드롭) ───────────────────────────────
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
        if (c or 0) == 0:                 # 무호가봉 드롭
            continue
        out.append((symbol, f"{d} {tm}", o, h, lo, c, fetched))
    out.sort(key=lambda x: x[1])
    return out


# ── 단일 조회 (count만 유효) ─────────────────────────────────────────
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
        # ★Close=중간값 필수(매수/매도/중간값 선택). 연장판이라 분봉 지원.
        opt = (f"Per=분,Cycle={config.INTRADAY_CYCLE},Close=중간값,sort=A,real=false,"
               f"Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true")
        ws.Range("A2").Formula = (
            f'=IMDH("{spec["kind"]}","{spec["ticker"]}",A3:{last_col}3,$B$1,$D$1,$F$1,"{opt}")')
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
        return list(FX_SYMBOLS.items())
    return [(n, FX_SYMBOLS[n]) for n in names]


def _floor_str():
    return dt.datetime.fromisoformat(config.INTRADAY_BACKFILL_START).strftime("%Y-%m-%d %H:%M:%S")


def _ingest(conn, sym, spec, count, floor, newer_than=None):
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
                print(f"  {sym:8} {spec['label']:14} {bars[0][1]} ~ {bars[-1][1]} | {len(bars)}봉/{days}일")
            else:
                print(f"  {sym:8} {spec['label']:14} 데이터 없음")
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
    """이력 없으면 백필(99999), 있으면 증분(last_ts 이후만). run.bat 진입점."""
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
                if raw_min is not None and raw_min > last:
                    print(f"    · 공백 큼(최소ts {raw_min} > {last}) → 전체콜 재수집")
                    n, f, l, _ = _ingest(conn, sym, spec, full, floor, newer_than=last)
            print(f"    = {sym}: +{n}봉  {f} → {l}" if n else f"    = {sym}: 신규 없음")
    finally:
        _kill_headless_excel()
    print("FX 동기화 완료:", coverage(conn))
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
