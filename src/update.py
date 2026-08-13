# -*- coding: utf-8 -*-
"""
update.py — 오케스트레이터 (시스템의 심장, 사양서 §4)

    python update.py

멱등: 몇 번 돌려도 결과 동일. 증분: 소스별 마지막 수집일+1 ~ 오늘.
최초 실행(원장 비어있음)만 과거 BACKFILL_YEARS 백필.
"""
import csv
import sys
from datetime import date, timedelta

# Windows 콘솔이 cp949여도 한글 리포트가 안 깨지게
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import config
import store
from collectors import ACTIVE_COLLECTORS
import sanity


# ── 기간 결정 ────────────────────────────────────────────────────────
def window_for(conn, source, today):
    """(start, end). 원장에 해당 소스가 있으면 (마지막일+1 − 겹침일)~오늘, 없으면 백필.
    ★겹침(INCREMENTAL_OVERLAP_DAYS): last_date_for_source는 소스 필드 전체의 MAX date라,
    먼저 공표되는 필드(cd91)가 그날을 채우면 늦게 공표되는 필드(y_ktb)가 영구 결번됨.
    최근 며칠을 매번 재조회(멱등 upsert)해 늦게 나온 값을 다음 실행에 채운다."""
    last = store.last_date_for_source(conn, source)
    if last:
        start = (date.fromisoformat(last) + timedelta(days=1)
                 - timedelta(days=config.INCREMENTAL_OVERLAP_DAYS))
    else:
        start = config.BACKFILL_START     # 소스가 그 전 데이터 없으면 빈 응답으로 자동 스킵
    return start, today


# 국채선물 최종거래일 = 결제월(3·6·9·12) 세 번째 화요일
def _third_tuesday(y, m):
    import calendar as _cal
    tues = [d for d in _cal.Calendar().itermonthdates(y, m)
            if d.month == m and d.weekday() == 1]
    return tues[2]


def _days_to_expiry(d):
    """d(date) 기준 최근월물 잔존일수 = 다음 분기월 세 번째 화요일까지 (당일 이후)."""
    from datetime import date as _date
    y = d.year
    for m in (3, 6, 9, 12, 15):
        yy, mm = (y, m) if m <= 12 else (y + 1, m - 12)
        exp = _third_tuesday(yy, mm)
        if exp >= d:
            return (exp - d).days
    return 0


# ── 파생 계산 (사양서 §4-6) ──────────────────────────────────────────
def recompute_derived(conn):
    fetched = store.now_iso()
    # 1) fx_cum = fx_net 누적합
    for net_field, cum_field in [("fx_net_ktb3", "fx_cum_ktb3"),
                                 ("fx_net_ktb10", "fx_cum_ktb10")]:
        series = store.field_series(conn, net_field)  # {date: value} 오름차순
        running, rows = 0.0, []
        for d in sorted(series):
            running += series[d]
            rows.append((d, cum_field, running, "derived", d, fetched))
        if rows:
            store.upsert(conn, rows)

    # 2) days_to_expiry = 분기 3번째 화요일까지 (거래일마다 날짜에서 계산)
    from datetime import date as _date
    dates = store.all_dates(conn)
    rows = [(d, "days_to_expiry", float(_days_to_expiry(_date.fromisoformat(d))),
             "derived", d, fetched) for d in dates]
    if rows:
        store.upsert(conn, rows)

    # 3) roll_flag = days_to_expiry가 전일 대비 상승하는 날(=최근월물 롤오버)
    dte = store.field_series(conn, "days_to_expiry")
    rows, prev = [], None
    for d in sorted(dte):
        flag = 1.0 if (prev is not None and dte[d] > prev) else 0.0
        rows.append((d, "roll_flag", flag, "derived", d, fetched))
        prev = dte[d]
    if rows:
        store.upsert(conn, rows)

    # 4) base_rate 전진보간: 정책금리는 금통위까지 불변 → ECOS 공표지연 꼬리를 채운다.
    #    (시장데이터와 달리 '모름'이 아니라 '같은 값 지속'이라 ffill이 정당)
    br = store.field_series(conn, "base_rate")
    if br:
        rows, last = [], None
        for d in sorted(set(store.all_dates(conn)) | set(br)):
            if d in br:
                last = br[d]
            elif last is not None:
                rows.append((d, "base_rate", last, "derived", d, fetched))
        if rows:
            store.upsert(conn, rows)


# ── 값 포맷 (멱등 export — 바이트 동일 보장) ─────────────────────────
_INT_FIELDS = (config.NONNEG_INT_FIELDS
               | {"days_to_expiry", "roll_flag"}
               | {f for f in config.DATA_FIELDS if f.endswith(("_net_ktb3", "_net_ktb10"))}
               | {"fx_cum_ktb3", "fx_cum_ktb10"})


def fmt(field, v):
    if v is None:
        return ""
    if field in _INT_FIELDS:
        return str(int(round(v)))
    s = ("%.6f" % v).rstrip("0").rstrip(".")
    return s if s not in ("", "-0") else "0"


def is_trading_day(fields_present):
    """거래일 = KTB 선물이 거래된 날(선물 필드가 하나라도 있는 날)."""
    return bool(set(fields_present) & config.TRADING_ANCHOR_FIELDS)


def export_csv(conn):
    table = store.wide_table(conn)
    # 휴장일 행 없음: 시장 필드가 없는 날(주말/공휴일에 base_rate만 있는 날)은 제외
    dates = sorted(d for d, fm in table.items() if is_trading_day(fm))
    config.OUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(config.OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(config.COLUMNS)              # 헤더 (42컬럼 고정순)
        for d in dates:
            rowmap = table[d]
            row = [d if c == "date" else "" if c == "note"
                   else fmt(c, rowmap.get(c)) for c in config.COLUMNS]
            w.writerow(row)
    return len(dates)


# ── 수기 입력 병합 (라이선스로 자동불가한 kr_cds5 등) ────────────────
def merge_manual(conn):
    """manual_input.csv (date,field,value) → ledger(source='manual').
    S&P 라이선스로 애드인 다운로드 막힌 kr_cds5 등 수기 입력용. 없으면 skip."""
    if not config.MANUAL_CSV.exists():
        return 0
    from collectors.base import Record
    recs = []
    with open(config.MANUAL_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            d = (row.get("date") or "").strip()
            fld = (row.get("field") or "").strip()
            v = (row.get("value") or "").strip()
            if not (d and fld and v) or fld not in config.DATA_FIELDS:
                continue
            recs.append(Record(date=d, field=fld, value=float(v), as_of=d))
    if not recs:
        return 0
    existing = {f: store.field_series(conn, f) for f in {r.field for r in recs}}
    passed, _flag, _drop = sanity.run(recs, existing)
    fetched = store.now_iso()
    store.upsert(conn, [(r.date, r.field, r.value, "manual", r.as_of, fetched)
                        for r in passed])
    return len(passed)


# ── 완결성 점검: 거래일인데 빠진 소스를 시끄럽게 (휴일 빈칸과 구분) ────
def completeness_report(conn, ref):
    """XKRX 거래일 대비 최근 누락 점검. 거래일 아닌 날(휴일)은 애초에 검사 안 함.
    맨 끝 1거래일은 공표 지연 유예. 반환: 경보 줄 리스트."""
    from datetime import timedelta
    sess = config.trading_days(ref - timedelta(days=25), ref)
    check = sess[:-1][-7:]                      # 최근 7거래일(맨 끝 유예)
    if not check:
        return []
    lines = []
    for src, fld in config.KEY_FIELD.items():
        have = set(store.field_series(conn, fld).keys())
        miss = [d.isoformat() for d in check if d.isoformat() not in have]
        if miss:
            lines.append(f"  ⚠ [완결성] {src}({fld}) 거래일 누락: {', '.join(miss)}")
    return lines


# ── 메인 ─────────────────────────────────────────────────────────────
def main():
    today = config.ref_date()      # 기준일 = 마지막 완료 영업일 (오늘 아님)
    conn = store.connect()
    store.init_db(conn)

    report = []           # 소스별 리포트 줄
    total_pass = total_flag = total_drop = 0

    # HTTP 소스 수집 → sanity → upsert (인포맥스는 infomax_pull.py로 별도 적재)
    for col in ACTIVE_COLLECTORS:
        source = col.source
        start, end = window_for(conn, source, today)
        last = store.last_date_for_source(conn, source)
        if start > end:
            report.append(f"{source:6}: 최종 {last} → 신규 없음  OK")
            continue
        try:
            records = col.fetch(start, end)
        except Exception as e:
            # 실패는 시끄럽게 — 삼키지 않는다
            report.append(f"{source:6}: 최종 {last} → FETCH 실패: {e}  ✗")
            continue

        # sanity: field별 직전값 기준
        prev = {fld: (store.field_series(conn, fld) or {})
                for fld in {r.field for r in records}}
        passed, flagged, dropped = sanity.run(records, prev)

        fetched = store.now_iso()
        store.upsert(conn, [(r.date, r.field, r.value, source, r.as_of, fetched)
                            for r in passed])
        # 플래그(동결)+드롭(격리) 둘 다 quarantine 테이블에 사유로 남긴다(가시성).
        logrows = ([(r.date, r.field, r.value, source, "FLAG:" + why, fetched)
                    for r, why in flagged]
                   + [(r.date, r.field, r.value, source, "DROP:" + why, fetched)
                      for r, why in dropped])
        if logrows:
            store.quarantine(conn, logrows)

        total_pass += len(passed)
        total_flag += len(flagged)
        total_drop += len(dropped)
        added_dates = len({r.date for r in passed})
        tags = []
        if flagged:
            tags.append(f"플래그 {len(flagged)}")
        if dropped:
            tags.append(f"격리 {len(dropped)}")
        tail = "  ⚠ " + ", ".join(tags) if tags else "  OK"
        report.append(f"{source:6}: 최종 {last} → +{added_dates}일 ({len(passed)}값){tail}")

    # 수기 입력 병합 (kr_cds5 등 라이선스로 자동불가)
    n_manual = merge_manual(conn)
    if n_manual:
        report.append(f"manual: manual_input.csv → +{n_manual}값  OK")

    # 6: 파생 계산 (fx_cum, roll_flag)
    recompute_derived(conn)
    # 7: export
    n_rows = export_csv(conn)

    # 8: 건강 리포트
    print(f"\n=== UPDATE 기준일 {today} (실행 {store.now_iso()[11:16]}) ===")
    for line in report:
        print(line)
    print(f"[SANITY] 적재 {total_pass} (플래그 {total_flag}) / 격리 {total_drop}")
    comp = completeness_report(conn, today)
    if comp:
        print("[완결성] 최근 거래일 누락 감지:")
        for line in comp:
            print(line)
    else:
        print("[완결성] 최근 7거래일 핵심소스 이상 없음  OK")
    print(f"[EXPORT] {config.OUT_CSV}  {n_rows}행")
    conn.close()


if __name__ == "__main__":
    sys.exit(main())
