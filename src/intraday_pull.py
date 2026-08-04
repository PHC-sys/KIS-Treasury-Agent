# -*- coding: utf-8 -*-
"""
intraday_pull.py — 인포맥스 IMDH 5분봉 수집기 (선물·현물).

일별 파이프라인(infomax_pull)과 별개. 저장은 db/intraday.sqlite (intraday_store).
증분 원칙: 매일 마지막 ts 이후만 저장(멱등 upsert). 백필은 최초 1회.

모드 (CLI):
  probe            종목별 가용 이력 깊이 확인 (최대콜 1번/종목, 저장 안 함)
  backfill         최초 이력 적재 (종목별 count=99999 한 콜 = 가용 전이력)
  sync (기본)      원클릭: 이력 없으면 백필, 있으면 증분(last_ts 이후만). intraday.bat

★ IMDH 분봉 gotchas (실측 확정):
  1) 계산에 시간 걸림 → 데이터 도착(A4값)을 폴링(계산상태 아님). 짧으면 0봉.
  2) ★B1/D1(날짜창)이 전부 무시됨 — 오직 count개가 '오늘 기준 뒤로' 온다.
     → 날짜 페이지네이션 불가. 가용 전이력 = count=99999 한 콜.
       (실측 선물 C65 2021-07-26~ / 현물 국고3년 2023-06-23~, 각 15~17초)
     세션시각·저장하한(floor)은 Python에서 필터.
  3) 현물 무체결봉(체결거래량=0, 값 carry)은 저장 시 드롭.
  4) ★무거운 콜을 한 앱으로 연달아 하면 Excel RPC 사망 → 종목마다 프레시 Excel.
  5) run.bat과 마찬가지로 infomax_data.xlsx는 닫아두고 실행(좀비 Excel 방지).
"""
import datetime as dt
import os
import sys
import time

import config
import intraday_store as store
from infomax_pull import (
    _register_addin, _com_retry, _kill_headless_excel,
)

XL_CALC_MANUAL = -4135
SESSION_START, SESSION_END = config.INTRADAY_SESSION


# ── COM 앱 수명주기 (검증된 sandbox/pull_5min_test 순서) ──────────────
def _open_app():
    """전용 hidden Excel을 새로(Dispatch) 띄운다 — 사용자 Excel은 안 건드림.
    ★순서 중요(sandbox 실증): blank 워크북 먼저 → manual 계산 → addin 등록.
    (워크북 0개 상태에서 Calculation 설정하면 안 먹혀 자동계산으로 IMDH가 백그라운드
     재계산→Excel busy→Workbooks.Add가 0x800ac472로 거부됨.)
    Dispatch 필수(DispatchEx는 애드인 안 붙어 #NAME?). 반환 (app, blank)."""
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    app = win32.Dispatch("Excel.Application")
    for _ in range(15):
        try:
            app.Version
            break
        except Exception:
            time.sleep(1)
    app.DisplayAlerts = False
    app.Visible = False
    blank = _com_retry(lambda: app.Workbooks.Add())   # ★ 계산모드 고정용 워크북 먼저
    try:
        app.Calculation = XL_CALC_MANUAL              # 이제(워크북 있음) 확실히 먹음
    except Exception:
        pass
    if not _register_addin(app):
        print("  ⚠ 인포맥스 XLL 등록 실패 → IMDH #NAME? 위험 (IMX_XLL 확인)")
    return app, blank


def _close_app(app, blank):
    try:
        blank.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


_ERR_MARKERS = ("#NAME?", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#DIV/0!", "#NULL!")


# ── 데이터 도착 폴링 (gotcha #1) ──────────────────────────────────────
def _wait_calc(app, ws, timeout=None, settle=3, poll=2, nudge=15, beat=15, tag=""):
    """IMDH가 서버에서 데이터를 실제로 채울 때까지 대기 (하트비트 + 에러 감지).
    ★ IMDH는 비동기 XLL — CalculationState가 xlDone이어도 데이터는 아직 안 왔을 수 있다.
       그래서 계산상태가 아니라 **첫 데이터셀(A4) 값 + UsedRange 행수**를 본다: A4에 값이
       들어오고 행수가 연속 settle회 불변이면 도착·안정 완료. 비어있는 동안 nudge로 재계산.
    - A2(수식셀)가 #NAME? 등 에러면 즉시 중단(애드인/코드 문제 — 기다려도 안 옴).
    - beat초마다 진행 로그(경과·행수·A2상태) 출력 → 깜깜이 대기 방지.
    반환 = UsedRange 행수(헤더 포함). 0이면 데이터 없음(에러/타임아웃)."""
    timeout = timeout or int(os.environ.get("IMX_INTRADAY_WAIT", "300"))
    _com_retry(lambda: app.CalculateFullRebuild())   # busy면 재시도(거부 크래시 방지)
    t0 = last_nudge = last_beat = time.time()
    prev, stable, a4 = -1, 0, None
    while time.time() - t0 < timeout:
        time.sleep(poll)
        try:
            a2 = ws.Cells(2, 1).Value            # 수식셀(첫 결과/에러 마커)
            a4 = ws.Cells(4, 1).Value            # 첫 데이터셀 — 값이 실제로 들어왔나
            nrows = ws.UsedRange.Rows.Count
        except Exception:
            a2, a4, nrows = None, None, prev
        el = time.time() - t0
        if isinstance(a2, str) and any(m in a2 for m in _ERR_MARKERS):    # 에러 → 즉시 중단
            print(f"      ⚠ {tag} IMDH 에러 '{a2.strip()[:30]}' ({el:.0f}s) — 애드인/코드 확인")
            return 0
        # 값 도착(A4 채워짐) AND 행수 불변 = 완료 (행범위 선할당 후 값 지연 채움 대비)
        if a4 is not None and nrows > 3 and nrows == prev:
            stable += 1
            if stable >= settle:
                print(f"      · {tag} 도착 {nrows - 3}행 ({el:.0f}s)")
                return nrows
        else:
            stable = 0
        prev = nrows
        if time.time() - last_beat >= beat:      # 하트비트
            st = "값대기" if a4 is None else f"채워짐(행 {nrows})"
            print(f"      … {tag} 대기 {el:.0f}s | UsedRange {nrows}행 | A2={str(a2)[:18]!r} | {st}")
            last_beat = time.time()
        if a4 is None and time.time() - last_nudge >= nudge:   # 아직 빈 채면 재계산 유도
            try:
                _com_retry(lambda: app.CalculateFullRebuild())
            except Exception:
                pass
            last_nudge = time.time()
    print(f"      ⚠ {tag} 타임아웃 {timeout}s — 데이터 안 옴 (IMX_INTRADAY_WAIT 상향 또는 애드인 확인)")
    return prev if (a4 is not None) else 0


# ── 파싱 헬퍼 ────────────────────────────────────────────────────────
def _fnum(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def _fdate(v):
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()[:10].replace("/", "-").replace(".", "-")
    return s if len(s) == 10 else None


def _ftime(v):
    """IMDH 시간 셀(TmFmt=1) → 'hh:mm:ss'. 문자열/시간객체/엑셀분수 모두 처리."""
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.strftime("%H:%M:%S")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, (int, float)):          # 엑셀 시간 분수(0~1)
        secs = round(float(v) % 1 * 86400)
        return f"{secs // 3600:02d}:{secs % 3600 // 60:02d}:{secs % 60:02d}"
    s = str(v).strip()
    return s[:8] if len(s) >= 8 else s


def _parse(rows, spec, symbol):
    """IMDH 결과 행렬 → bar 튜플 리스트(오름차순). 세션시각 필터 + 현물 무체결봉 드롭.
    항목 순서: 일자,시간,시가,고가,저가,현재가,(거래량|체결거래량),[선물만 미결제약정수량]."""
    is_fut = spec["kind"] == "FUT"
    fetched = store.now_iso()
    out = []
    for r in rows:
        if not r or r[0] is None:
            continue
        d = _fdate(r[0])
        tm = _ftime(r[1]) if len(r) > 1 else None
        if not d or not tm:
            continue
        if not (SESSION_START <= tm <= SESSION_END):    # 세션 밖(장전/장후) 드롭
            continue
        o, h, lo, c = (_fnum(r[i]) if i < len(r) else None for i in (2, 3, 4, 5))
        if o is None and h is None and lo is None and c is None:
            continue                                    # 완전 결측 행
        vol = _fnum(r[6]) if len(r) > 6 else None
        oi = _fnum(r[7]) if (is_fut and len(r) > 7) else None   # 선물만: 미결제약정수량
        if not is_fut and (vol or 0) == 0:              # 현물 무체결봉(carry) 드롭
            continue
        ts = f"{d} {tm}"
        out.append((symbol, ts, o, h, lo, c, vol, oi, fetched))
    out.sort(key=lambda x: x[1])
    return out


# ── 단일 조회 (날짜창 무의미 → count만 유효) ──────────────────────────
def _pull(app, symbol, spec, count):
    """symbol의 최근 count봉을 조회 → 파싱된 bar 리스트(오름차순).
    ★B1/D1은 무시되고 count개가 오늘 기준 뒤로 옴 → 가용 전이력 = count=99999."""
    wb = _com_retry(lambda: app.Workbooks.Add())   # busy(0x800ac472)면 재시도
    ws = wb.Worksheets(1)
    try:
        items = spec["items"]
        ncols = len(items)
        # B1/D1은 무시되지만 형식상 채움(B1=먼과거, D1=오늘)
        ws.Range("B1").Value = dt.datetime(2000, 1, 1)
        ws.Range("D1").Value = dt.datetime.now()
        ws.Range("F1").Value = count
        for j, it in enumerate(items):
            ws.Cells(3, j + 1).Value = it
        last_col = chr(ord("A") + ncols - 1)
        opt = (f"Per={spec['per']},Cycle={config.INTRADAY_CYCLE},sort=A,"
               f"real={spec['real']},Bizday=0,Quote=종가,Pos=20,Orient=V,"
               f"Title=T,DtFmt=1,TmFmt=1,unit=true")
        ws.Range("A2").Formula = (
            f'=IMDH("{spec["kind"]}","{symbol}",A3:{last_col}3,$B$1,$D$1,$F$1,"{opt}")')
        n = _wait_calc(app, ws, tag=symbol)
        if n <= 3:                                  # 에러/타임아웃 → 빈 결과
            return []
        used = ws.UsedRange.Rows.Count
        if used < 4:
            return []
        rng = ws.Range(ws.Cells(4, 1), ws.Cells(used, ncols)).Value
        rows = rng if isinstance(rng, tuple) else ((rng,),)
        return _parse(rows, spec, symbol)
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


# ── 공개 진입점 ──────────────────────────────────────────────────────
def _symbols(names):
    if not names:
        return list(config.INTRADAY_SYMBOLS.items())
    return [(n, config.INTRADAY_SYMBOLS[n]) for n in names]


def _floor_str():
    return dt.datetime.fromisoformat(config.INTRADAY_BACKFILL_START).strftime("%Y-%m-%d %H:%M:%S")


def probe(names=None):
    """종목별 가용 이력 깊이 확인 (최대콜 1번). 저장 안 함. 종목마다 프레시 앱."""
    for sym, spec in _symbols(names):
        _kill_headless_excel()
        app, blank = _open_app()
        try:
            bars = _pull(app, sym, spec, config.INTRADAY_BACKFILL_COUNT)
        finally:
            _close_app(app, blank)
        if bars:
            days = len({b[1][:10] for b in bars})
            print(f"  {sym:14} {spec['label']:12} {bars[0][1]} ~ {bars[-1][1]} "
                  f"| {len(bars)}봉 / {days}거래일")
        else:
            print(f"  {sym:14} {spec['label']:12} 데이터 없음")
    _kill_headless_excel()


def _ingest(conn, sym, spec, count, floor, newer_than=None):
    """종목 1개: 프레시 앱으로 count봉 조회 → floor/증분 필터 → upsert. 반환 (n, first, last)."""
    _kill_headless_excel()
    app, blank = _open_app()
    try:
        bars = _pull(app, sym, spec, count)
    finally:
        _close_app(app, blank)
    if not bars:
        return 0, None, None, None
    raw_min = bars[0][1]
    bars = [b for b in bars if b[1] >= floor]                 # 저장 하한
    if newer_than is not None:                                # 증분: last_ts 이후만
        bars = [b for b in bars if b[1] > newer_than]
    if bars:
        store.upsert(conn, bars)
    return len(bars), (bars[0][1] if bars else None), (bars[-1][1] if bars else None), raw_min


def backfill(names=None):
    """최초 이력 적재: 종목별 count=99999 한 콜 = 가용 전이력. 종목마다 프레시 앱."""
    conn = store.connect()
    store.init_db(conn)
    floor = _floor_str()
    for sym, spec in _symbols(names):
        print(f"  ▶ backfill {sym} ({spec['label']})")
        n, f, l, _ = _ingest(conn, sym, spec, config.INTRADAY_BACKFILL_COUNT, floor)
        print(f"    = {sym}: {n}봉  {f} → {l}" if n else f"    = {sym}: 데이터 없음")
    print("backfill 완료:", store.coverage(conn))
    conn.close()


def sync(names=None):
    """원클릭 갱신: 종목별로 이력 없으면 백필(99999), 있으면 증분(last_ts 이후만).
    증분은 작은 count로 뜨고, 공백이 커서 count가 last_ts까지 못 미치면 99999로 승격.
    intraday.bat 진입점 — 첫 실행=전체백필, 이후=가벼운 증분으로 자동 분기."""
    conn = store.connect()
    store.init_db(conn)
    floor = _floor_str()
    incr = config.INTRADAY_INCR_COUNT
    full = config.INTRADAY_BACKFILL_COUNT
    for sym, spec in _symbols(names):
        last = store.last_ts(conn, sym)
        if last is None:                             # 이력 없음 → 백필
            print(f"  ▶ {sym} ({spec['label']}) 이력 없음 → 백필")
            n, f, l, _ = _ingest(conn, sym, spec, full, floor)
        else:                                        # 증분: 최근 incr봉 중 last_ts 이후만
            print(f"  ▶ {sym} ({spec['label']}) 증분 (마지막 {last})")
            n, f, l, raw_min = _ingest(conn, sym, spec, incr, floor, newer_than=last)
            if raw_min is not None and raw_min > last:   # incr가 공백을 못 덮음 → 99999로 승격
                print(f"    · 공백 큼(최소ts {raw_min} > {last}) → 전체콜로 재수집")
                n, f, l, _ = _ingest(conn, sym, spec, full, floor, newer_than=last)
        print(f"    = {sym}: +{n}봉  {f} → {l}" if n else f"    = {sym}: 신규 없음")
    print("동기화 완료:", store.coverage(conn))
    conn.close()


# 하위호환 별칭
incremental = sync


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "sync"
    names = sys.argv[2:] or None
    if mode in ("sync", "incremental"):
        sync(names)                       # 원클릭: 첫 실행=백필, 이후=증분 (intraday.bat)
    elif mode == "probe":
        probe(names)
    elif mode == "backfill":
        backfill(names)
    else:
        print(__doc__)
