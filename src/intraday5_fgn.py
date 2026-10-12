# -*- coding: utf-8 -*-
"""
intraday5_fgn.py — 외국인 순매수(금액·수량) 5분봉 캡처 (국채선물 C65/C67).

★한계(실측 2026-10-12): 인포맥스 5분봉 외국인순매수는 **당일만 라이브로 제공**,
  과거는 전부 0(미저장). → 과거 백필 불가. **매 거래일 오후(마감 후) 캡처해 누적**한다.
  놓친 거래일은 영구 결번(백필 수단 없음).

★오후 run에서만 동작: now.hour < config.CLOSE_HOUR(16) 이면 건너뛴다(장중/아침엔 미완성).
  'force' 인자로 게이트 우회(테스트용).

저장: 기존 bars 테이블(intraday.sqlite)에 컬럼 2개(fgn_net_amt, fgn_net_qty)를 추가해
  당일 C65/C67 봉의 해당 컬럼만 upsert. OHLC 수집기(intraday_pull/intraday_store)는 무수정
  — INSERT가 명시 컬럼만 쓰고 ON CONFLICT도 OHLC 컬럼만 갱신하므로 fgn 컬럼은 보존된다.

레시피(참고 외국인순매수_5분.xlsx): FUT/C65·C67, 항목 외국인합순매수금액/수량, Per=분/Cycle=5, real=true.
  값은 **장중 누적(running total)** — 15:45 봉이 그날 총 외국인순매수(≈일별 fx_net).
"""
import datetime as dt
import os
import sqlite3
import sys

import config
import intraday_store as store                 # bars 테이블 보장용(무수정 재사용)
from intraday_pull import (
    _open_app, _close_app, _wait_calc, _kill_headless_excel, _com_retry,
    _fnum, _fdate, _ftime,
)

FGN_SYMBOLS = {"C65": "3년국채 연결", "C67": "10년국채 연결"}
ITEMS = ["일자", "시간", "외국인합순매수금액", "외국인합순매수수량"]
SESSION_START, SESSION_END = config.INTRADAY_SESSION
FGN_COLS = ("fgn_net_amt", "fgn_net_qty")


def connect():
    config.INTRADAY_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(config.INTRADAY_DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def init_db(conn):
    store.init_db(conn)                         # bars 없으면 생성(기존 스키마)
    have = {r[1] for r in conn.execute("PRAGMA table_info(bars)")}
    for c in FGN_COLS:                          # sqlite ALTER엔 IF NOT EXISTS 없음 → 수동 체크
        if c not in have:
            conn.execute(f"ALTER TABLE bars ADD COLUMN {c} REAL")
    conn.commit()


def upsert_fgn(conn, rows):
    """rows: (symbol, ts, amt, qty). **기존 OHLC 봉에 fgn 컬럼만 UPDATE**.
    ★INSERT가 아니라 UPDATE — fgn 봉이 OHLC 봉셋과 어긋나도(마감 직전 real 차이 등)
      OHLC 없는 orphan 행을 만들지 않는다. OHLC 수집이 run.bat에서 먼저 돌아 봉이 있으므로,
      매칭되는 봉만 fgn이 채워진다(없는 ts는 무시 = no-op)."""
    conn.executemany(
        "UPDATE bars SET fgn_net_amt=?, fgn_net_qty=? WHERE symbol=? AND ts=?",
        [(amt, qty, sym, ts) for (sym, ts, amt, qty) in rows])
    conn.commit()


def _pull_today(app, symbol, today, count=1000):
    """symbol의 최근 count봉 중 **오늘(today) 세션 봉**만 (symbol, ts, amt, qty)로.
    외국인순매수는 당일만 실값이라 오늘 날짜로 필터하면 그날치만 남는다."""
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
        opt = (f"Per=분,Cycle={config.INTRADAY_CYCLE},sort=A,real=true,Bizday=0,"
               f"Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true")
        ws.Range("A2").Formula = (
            f'=IMDH("FUT","{symbol}",A3:{last_col}3,$B$1,$D$1,$F$1,"{opt}")')
        n = _wait_calc(app, ws, tag=f"fgn-{symbol}")
        if n <= 3:
            return []
        used = ws.UsedRange.Rows.Count
        rng = ws.Range(ws.Cells(4, 1), ws.Cells(used, ncols)).Value
        rows = rng if isinstance(rng, tuple) else ((rng,),)
        out = []
        for r in rows:
            if not r or r[0] is None:
                continue
            d = _fdate(r[0])
            tm = _ftime(r[1]) if len(r) > 1 else None
            if d != today or not tm:
                continue
            if not (SESSION_START <= tm <= SESSION_END):
                continue
            amt = _fnum(r[2]) if len(r) > 2 else None
            qty = _fnum(r[3]) if len(r) > 3 else None
            if qty is None and amt is None:
                continue
            out.append((symbol, f"{d} {tm}", amt, qty))
        out.sort(key=lambda x: x[1])
        return out
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def capture_one(symbol):
    """단일 심볼 캡처(프로세스당 1종목 = COM 격리). capture()가 서브프로세스로 호출.
    ★한 프로세스에서 C65·C67 연속 COM 호출은 두 번째에서 Excel이 죽는다(실측). 반드시 분리."""
    today = dt.date.today().isoformat()
    conn = connect()
    init_db(conn)
    try:
        _kill_headless_excel()
        app, blank = _open_app()
        try:
            bars = _pull_today(app, symbol, today)
        finally:
            _close_app(app, blank)
            _kill_headless_excel()
        if bars:
            upsert_fgn(conn, bars)
            print(f"  ▶ {symbol} ({FGN_SYMBOLS.get(symbol, '')}) 외국인순매수 {len(bars)}봉 "
                  f"({bars[0][1][11:16]}~{bars[-1][1][11:16]}) 막판 qty={bars[-1][3]}")
        else:
            print(f"  ▶ {symbol} ({FGN_SYMBOLS.get(symbol, '')}) 당일 봉 없음(휴장/미집계?)")
    finally:
        conn.close()
    return len(bars)


def capture(force=False):
    """오후(마감 후) 당일 외국인순매수 캡처. run.bat에서 호출(인자 없음=게이트 적용).
    심볼마다 별도 파이썬 프로세스로 돌려 COM을 격리한다(2종목 연속 크래시 방지)."""
    # ★스키마(fgn 컬럼)는 게이트와 무관하게 항상 보장한다 — push_intraday5가 이 컬럼을
    #   SELECT하므로, 아침 run에서 게이트로 캡처를 건너뛰더라도 컬럼은 존재해야 한다.
    _c = connect()
    init_db(_c)
    _c.close()
    now = dt.datetime.now()
    if not force and now.hour < config.CLOSE_HOUR:
        print(f"  외국인순매수: 장 마감 전({now.hour:02d}시 < {config.CLOSE_HOUR}시) — "
              f"오후 run에서만 캡처. 건너뜀")
        return
    today = now.date().isoformat()
    try:
        tds = {d.isoformat() for d in config.trading_days(
            now.date() - dt.timedelta(days=10), now.date())}
        if today not in tds:
            print(f"  외국인순매수: 오늘({today}) 거래일 아님 — 건너뜀")
            return
    except Exception:
        pass
    import subprocess
    this = os.path.abspath(__file__)
    ok = 0
    for sym in FGN_SYMBOLS:
        r = subprocess.run([sys.executable, this, "_one", sym])   # COM 격리(심볼당 프로세스)
        if r.returncode == 0:
            ok += 1
        else:
            print(f"  ⚠ {sym} 캡처 실패(returncode {r.returncode}) — 로그인/COM 확인")
    print(f"외국인순매수 캡처 완료: {ok}/{len(FGN_SYMBOLS)} 종목 ({today})")


if __name__ == "__main__":
    _args = sys.argv[1:]
    if _args and _args[0] == "_one":              # 서브프로세스 워커(게이트 없음)
        capture_one(_args[1])
    else:
        capture(force=("force" in _args))
