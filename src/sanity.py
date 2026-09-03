# -*- coding: utf-8 -*-
"""
sanity.py — 적재 전 오염 방지 (사양서 §5)

하나라도 실패하면 그 (date, field)만 격리 + 로그. 조용히 넘어가지 않는다.
sanity 실패는 에러가 아니라 경고+격리 — 나머지는 계속 진행.
"""
import config


def check_record(rec, prev_value):
    """Record 하나 검증. 반환 (reason, hard).
      reason=None      → 통과
      hard=True(드롭)  → 값이 물리적으로 틀림(부호·형식·과대점프). 적재 중단+격리.
      hard=False(플래그) → 의심되나 값은 살림(동결). 적재하되 로그에 남김(사양서 §5).
    prev_value: 같은 field의 직전 거래일 값 (없으면 None)."""
    f, v = rec.field, rec.value

    # 형식: 결측(None)은 애초에 Record가 되면 안 됨
    if v is None:
        return "value_is_none", True

    # 부호/형식: vol·oi ≥ 0 정수
    if f in config.NONNEG_INT_FIELDS:
        if v < 0:
            return f"negative_count({v})", True
        if float(v) != int(v):
            return f"non_integer_count({v})", True

    # 비현실적 0: 최근월물 거래량·미결제가 0 = 피드오류 → 격리(가짜 0 저장 방지).
    # 수급(net)의 0은 정상이라 여기 안 걸림.
    if f in config.ZERO_IMPLAUSIBLE and v == 0:
        return "implausible_zero", True

    # 양수여야 하는 금리류 (>0). 순매수(net)는 음수 정상이라 제외.
    if f in config.POSITIVE_FIELDS and v <= 0:
        return f"non_positive({v})", True

    # 물리적 범위: 전일 대비 과대변화 → **플래그(적재하되 경고)**, 드롭 아님.
    #  이유: 실제 급변(위기·금리인하)·연결선물 롤갭은 진짜 값이라 버리면 안 됨.
    #  드롭하면 다음날이 옛 기준과 비교돼 연쇄격리(cascade)됨. 그로스 오류만 걸러 플래그.
    if prev_value is not None:
        if f.startswith("usdkrw"):        # usdkrw(매매기준율) + usdkrw_open/high/low/close(OHLC)
            if prev_value and abs(v - prev_value) / abs(prev_value) > config.USDKRW_PCT_LIMIT:
                return f"jump_pct({(v-prev_value)/prev_value:.3%})", False
        elif f in config.RANGE_LIMITS:
            if abs(v - prev_value) > config.RANGE_LIMITS[f]:
                return f"jump_abs({v-prev_value:+.3f}>{config.RANGE_LIMITS[f]})", False

        # 동결 감지: 라이브 시리즈가 전일과 완전 동일 → 플래그+적재.
        if f in config.FREEZE_CHECK_FIELDS and v == prev_value:
            return "frozen(same_as_prev)", False

    return None, False


def run(records, existing):
    """records 전체를 검증.
    existing: {field: {date: value}} — 기존 ledger 전체 시계열.
    반환: (passed, flagged, dropped)
      passed   = 적재할 Record (통과 + 플래그 포함)
      flagged  = (rec, reason) — 적재하되 의심표시(동결 등)
      dropped  = (rec, reason) — 격리(물리적 오류)

    핵심: 각 record의 '직전값'은 **날짜 인접**(그 날짜보다 이전의 가장 최근 값)으로 찾는다.
    (과거 백필 시 최신값과 비교해 오탐 격리하던 버그 수정.)"""
    passed, flagged, dropped = [], [], []
    # field별 병합 시계열: 기존값 + 통과한 신규값
    timeline = {f: dict(existing.get(f, {}))
                for f in {r.field for r in records}}
    for rec in sorted(records, key=lambda r: (r.field, r.date)):
        tl = timeline[rec.field]
        earlier = [d for d in tl if d < rec.date]
        prev = tl[max(earlier)] if earlier else None
        reason, hard = check_record(rec, prev)
        if reason is not None and hard:
            dropped.append((rec, reason))
            continue
        passed.append(rec)
        tl[rec.date] = rec.value          # 통과값을 시계열에 반영
        if reason is not None:            # 플래그(동결) — 살리되 기록
            flagged.append((rec, reason))
    return passed, flagged, dropped
