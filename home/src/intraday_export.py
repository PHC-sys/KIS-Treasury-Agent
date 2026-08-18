# -*- coding: utf-8 -*-
r"""
intraday_export.py — db/intraday.sqlite(bars) → CSV 내보내기 (사람이 보기용).

사용:
  python src\intraday_export.py                 # 6종목 각각 전체 → out/intraday/<symbol>.csv
  python src\intraday_export.py C65             # C65 전체
  python src\intraday_export.py C65 --days 5    # C65 최근 5일
  python src\intraday_export.py --days 10       # 6종목 최근 10일
저장은 안 건드림(읽기 전용). 엑셀에서 바로 열림(UTF-8 BOM).
"""
import csv
import sys
from pathlib import Path

import config
import intraday_store as store

OUT_DIR = config.ROOT / "out" / "intraday"
COLS = ["symbol", "ts", "open", "high", "low", "close", "volume", "oi"]


def export(symbol, days=None):
    conn = store.connect()
    if days:
        # 그 종목 마지막 날짜에서 days일 전부터
        last = conn.execute("SELECT MAX(ts) FROM bars WHERE symbol=?", (symbol,)).fetchone()[0]
        if not last:
            print(f"  {symbol}: 데이터 없음"); conn.close(); return None
        import datetime as dt
        cutoff = (dt.date.fromisoformat(last[:10]) - dt.timedelta(days=days)).isoformat()
        rows = conn.execute(
            "SELECT symbol,ts,open,high,low,close,volume,oi FROM bars "
            "WHERE symbol=? AND ts>=? ORDER BY ts", (symbol, cutoff)).fetchall()
    else:
        rows = conn.execute(
            "SELECT symbol,ts,open,high,low,close,volume,oi FROM bars "
            "WHERE symbol=? ORDER BY ts", (symbol,)).fetchall()
    conn.close()
    if not rows:
        print(f"  {symbol}: 해당 구간 데이터 없음"); return None
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    suffix = f"_last{days}d" if days else ""
    path = OUT_DIR / f"{symbol}{suffix}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # BOM = 엑셀 한글/정렬 안전
        w = csv.writer(f)
        w.writerow(COLS)
        w.writerows(rows)
    print(f"  {symbol:14} {len(rows):>6}행  {rows[0][1]} ~ {rows[-1][1]}  → {path}")
    return path


def main(argv):
    days = None
    if "--days" in argv:
        i = argv.index("--days")
        days = int(argv[i + 1])
        argv = argv[:i] + argv[i + 2:]
    names = argv or list(config.INTRADAY_SYMBOLS)
    print(f"[export] {'최근 %d일' % days if days else '전체'} → {OUT_DIR}")
    for sym in names:
        export(sym, days)


if __name__ == "__main__":
    main(sys.argv[1:])
