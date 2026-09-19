# -*- coding: utf-8 -*-
"""push_intraday5.py — 5분봉 로컬(intraday.sqlite) → Supabase 2개 테이블 upsert.

★기존 push_supabase.py(일별) 무수정. 5분봉 전용 신규 push.
  - ktb_5min : 국채선물 C65/C67 (OHLC+거래량+OI)  ← intraday.sqlite bars
  - ust_5min : 미국채 US02Y/US10Y (OHLC only)      ← intraday.sqlite ust_bars

연결: 환경변수 SUPABASE_DB_URL (일별과 동일). 멱등: (symbol, ts) 충돌 시 덮어씀.
증분 push: 테이블별 심볼별 Supabase MAX(ts) 이후 로컬행만 전송(최초엔 전량).
"""
import os
import sqlite3

import psycopg2
from psycopg2.extras import execute_values

import config

# (Supabase 테이블, 로컬 테이블, 심볼필터, 컬럼) — 컬럼 순서 = 로컬 SELECT 순서
TABLES = {
    "ktb_5min": {
        "src": "bars",
        "symbols": ["C65", "C67"],
        "cols": ["symbol", "ts", "open", "high", "low", "close", "volume", "oi"],
    },
    "ust_5min": {
        "src": "ust_bars",
        "symbols": ["US02Y", "US10Y"],
        "cols": ["symbol", "ts", "open", "high", "low", "close"],
    },
    "fx_5min": {
        "src": "fx_bars",
        "symbols": ["USDKRW"],
        "cols": ["symbol", "ts", "open", "high", "low", "close"],
    },
}


def _coltype(c):
    if c == "symbol":
        return "text"
    if c == "ts":
        return "timestamp"
    return "double precision"


def _ensure_table(cur, table, cols):
    cols_sql = ",\n".join(f"  {c} {_coltype(c)}" for c in cols)
    cur.execute(f"CREATE TABLE IF NOT EXISTS {table} (\n{cols_sql},\n"
                f"  PRIMARY KEY (symbol, ts)\n)")
    for c in cols:                      # 신규 컬럼 자동 반영(있으면 무시)
        if c in ("symbol", "ts"):
            continue
        cur.execute(f"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {c} double precision")
    # 공개 시세데이터 — RLS 끔(anon엔 grant 안 함으로 통제, 일별과 동일 정책)
    cur.execute(f"ALTER TABLE {table} DISABLE ROW LEVEL SECURITY")


def _push_table(pg, lite, table, spec):
    cur = pg.cursor()
    _ensure_table(cur, table, spec["cols"])
    pg.commit()
    collist = ",".join(spec["cols"])
    updates = ",".join(f"{c}=EXCLUDED.{c}" for c in spec["cols"] if c not in ("symbol", "ts"))
    sql = (f"INSERT INTO {table} ({collist}) VALUES %s "
           f"ON CONFLICT (symbol, ts) DO UPDATE SET {updates}")
    total = 0
    for sym in spec["symbols"]:
        cur.execute(f"SELECT to_char(MAX(ts),'YYYY-MM-DD HH24:MI:SS') FROM {table} WHERE symbol=%s",
                    (sym,))
        maxts = cur.fetchone()[0]                     # None이면 전량
        q = f"SELECT {collist} FROM {spec['src']} WHERE symbol=?"
        args = [sym]
        if maxts:
            q += " AND ts > ?"
            args.append(maxts)
        q += " ORDER BY ts"
        data = lite.execute(q, args).fetchall()
        if data:
            execute_values(cur, sql, data, page_size=2000)
            pg.commit()
        print(f"    {table}.{sym}: +{len(data)}행 (기존 max {maxts or '없음'})")
        total += len(data)
    return total


def main():
    url = os.environ.get("SUPABASE_DB_URL")
    if not url:
        raise RuntimeError("SUPABASE_DB_URL 미설정")
    lite = sqlite3.connect(config.INTRADAY_DB_PATH)
    pg = psycopg2.connect(url)
    try:
        for table, spec in TABLES.items():
            print(f"  ▶ push {table} ← {spec['src']} {spec['symbols']}")
            _push_table(pg, lite, table, spec)
        # 최종 커버리지
        cur = pg.cursor()
        for table in TABLES:
            cur.execute(f"SELECT symbol, count(*), min(ts), max(ts) FROM {table} GROUP BY symbol ORDER BY symbol")
            for sym, n, lo, hi in cur.fetchall():
                print(f"    [SUPABASE] {table}.{sym}: {n}행 ({lo} ~ {hi})")
    finally:
        pg.close()
        lite.close()


if __name__ == "__main__":
    main()
