# -*- coding: utf-8 -*-
"""push_supabase.py — out/ktb_daily.csv → Supabase Postgres (ktb_daily) upsert.

연결은 환경변수 SUPABASE_DB_URL (Session pooler URI). 멱등: (date) 충돌 시 덮어쓰기.
매일 run.py 끝에 붙여 자동 반영.
"""
import csv
import os

import psycopg2
from psycopg2.extras import execute_values

import config


def _coltype(c):
    if c == "date":
        return "date primary key"
    if c == "note":
        return "text"
    return "double precision"


def _conv(c, v):
    if v is None or v == "":
        return None
    if c in ("date", "note"):
        return v
    return float(v)


def main():
    url = os.environ.get("SUPABASE_DB_URL")
    if not url:
        raise RuntimeError("SUPABASE_DB_URL 미설정")
    conn = psycopg2.connect(url)
    cur = conn.cursor()

    # 테이블 생성(없으면)
    cols_sql = ",\n".join(f"  {c} {_coltype(c)}" for c in config.COLUMNS)
    cur.execute(f"CREATE TABLE IF NOT EXISTS ktb_daily (\n{cols_sql}\n)")
    # RLS 끔: 공개 시세데이터, 접근은 role GRANT로 통제(anon엔 grant 안 함).
    cur.execute("ALTER TABLE ktb_daily DISABLE ROW LEVEL SECURITY")
    conn.commit()

    # CSV 읽어 upsert
    with open(config.OUT_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    data = [[_conv(c, r.get(c)) for c in config.COLUMNS] for r in rows]
    collist = ",".join(config.COLUMNS)
    updates = ",".join(f"{c}=EXCLUDED.{c}" for c in config.COLUMNS if c != "date")
    sql = (f"INSERT INTO ktb_daily ({collist}) VALUES %s "
           f"ON CONFLICT (date) DO UPDATE SET {updates}")
    execute_values(cur, sql, data, page_size=1000)
    conn.commit()

    cur.execute("SELECT count(*), min(date), max(date) FROM ktb_daily")
    n, lo, hi = cur.fetchone()
    print(f"[SUPABASE] push 완료: {n}행 ({lo} ~ {hi})")
    conn.close()


if __name__ == "__main__":
    main()
