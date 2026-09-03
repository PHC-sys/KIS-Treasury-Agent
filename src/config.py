# -*- coding: utf-8 -*-
"""
config.py — 시스템 전역 상수 (스키마·소스경계·sanity 임계값)

여기가 42컬럼 스키마와 "어느 필드를 어느 소스가 채우나"의 단일 진실.
스키마는 차장님 지정 — 순서·이름 변경 금지 (사양서 §2).
"""
from datetime import datetime, timedelta
from pathlib import Path

# ── 경로 ────────────────────────────────────────────────────────────
# 코드는 src/ 아래, 데이터(db·raw·out·manual)는 프로젝트 루트 아래.
ROOT = Path(__file__).resolve().parent.parent   # = 프로젝트 루트 (src/의 부모)
DB_PATH = ROOT / "db" / "ktb.sqlite"
INTRADAY_DB_PATH = ROOT / "db" / "intraday.sqlite"   # 5분봉 (선물·현물) — 별도 파일·증분 전용
RAW_DIR = ROOT / "raw"
OUT_DIR = ROOT / "out"
OUT_CSV = OUT_DIR / "ktb_daily.csv"
MANUAL_CSV = ROOT / "manual_input.csv"   # 인포맥스 4개 병합 경로 (비어 있어도 됨)

# ── 42컬럼 (순서 고정 — 사양서 §2) ───────────────────────────────────
# date는 행 키. note는 사람이 채움(수집기는 비움).
COLUMNS = [
    "date", "ktb3_settle", "ktb3_chg", "ktb10_settle", "ktb10_chg",
    "ktb3_vol", "ktb10_vol", "ktb3_oi", "ktb10_oi",
    "fx_net_ktb3", "fx_net_ktb10", "fx_cum_ktb3", "fx_cum_ktb10",
    "bank_net_ktb3", "itrust_net_ktb3", "ins_net_ktb3",
    "y_ktb3", "y_ktb10", "y_ktb30",
    "base_rate", "cd91", "kofr", "irs3y", "ois3y", "ust10", "usdkrw", "kr_cds5",
    "ktb3_theo_basis", "ktb10_theo_basis", "note",
    "ktb3_open", "ktb3_high", "ktb3_low",
    "ktb10_open", "ktb10_high", "ktb10_low",
    "days_to_expiry", "roll_flag",
    "bank_net_ktb10", "itrust_net_ktb10", "ins_net_ktb10", "irs10y",
    # v2.5 MACRO 축 (인포맥스 ECO, 발표일 배치)
    "cpi_yoy", "core_cpi_yoy", "bei", "exports_yoy", "ip_yoy",
    "wti",   # WTI 유가 종가(인포맥스 FRN/SPT:CL 현재가). 차장님 지시로 추가. as-of 시프트(미국일→한국).
    # ── v2.6 수급 세분화 (차장님 요청) ─────────────────────────────────
    # 인포맥스 FUT 투자자별 '순매수수량'. 단위=계약, 부호 +매수/−매도 (기존 수급과 동일).
    #  etc_net = "기타법인·국가 등 나머지 전부의 합" — 차장님 정의 그대로
    #            = 기타 + 국가/지방 + 연기금 + 종신금.
    #  → 항등식 fx+bank+itrust+ins+sec+indiv+etc = 0 으로 적재 정합성 매일 자동검증(FLOW_IDENTITY).
    #  pension_net은 **etc_net의 부분집합**(독립 컬럼 아님). 실측상 etc의 86~88%가 연기금이라,
    #  '기타법인 흐름'으로 오독되지 않도록 따로 노출한다.
    #    ※ 8개를 전부 더하면 연기금이 이중계산된다 — 합산할 땐 pension 제외.
    "sec_net_ktb3", "sec_net_ktb10",
    "indiv_net_ktb3", "indiv_net_ktb10",
    "etc_net_ktb3", "etc_net_ktb10",
    "pension_net_ktb3", "pension_net_ktb10",
    # ── v2.7 미국채 2Y·10Y OHLC (차장님 요청) ─────────────────────────
    # 인포맥스 IR/US10Y·US02Y의 MID_Open/High/Low/Close. 단위 %, 소수 그대로.
    #  ★MID 기준 — 기존 ust10(26년치)이 MID_Close라 시리즈를 통일한다.
    #    단말 화면(FRN/GVO:TR10Y·GVO:TR02Y '현재가')은 BID이며 MID보다 0.0~0.5bp 높다.
    #    같은 IR 호출에 BID_*도 있으므로 필요하면 나중에 별도 컬럼으로 추가할 것(기존 시리즈 변경 금지).
    #  ust10은 기존 컬럼 그대로 두고(=ust10_close 역할) open/high/low만 추가한다.
    #  as-of 시프트(미국일→다음 한국거래일)는 ust10과 동일하게 적용.
    "ust10_open", "ust10_high", "ust10_low",
    "ust2", "ust2_open", "ust2_high", "ust2_low",
    # ── v2.8 USDKRW 정규장 OHLC (차장님 요청) ─────────────────────────
    # 인포맥스 FX/USDSP_SMBCC(정규 09:00~15:30) 시가/고가/저가/현재가. 한국거래라 as-of 시프트 없음.
    #  기존 usdkrw는 그대로 유지(=ECOS 매매기준율=당일 최초고시). 이건 시장 OHLC.
    #  close = 현재가(15:30 마감). 파이프라인이 채권종가(3:30) 후 실행이라 당일 반영됨.
    #  ★_EXT(연장/야간)가 아니라 정규 USDSP_SMBCC 사용.
    "usdkrw_open", "usdkrw_high", "usdkrw_low", "usdkrw_close",
]
assert len(COLUMNS) == 67, f"컬럼 수 {len(COLUMNS)} != 67"

# 수집 대상 필드 = date/note 제외 (ledger에 들어갈 것들)
DATA_FIELDS = [c for c in COLUMNS if c not in ("date", "note")]

# 파이프라인이 계산 (수집 아님):
#  fx_cum = fx_net 누적합 / days_to_expiry = 분기 3번째 화요일까지 / roll_flag = 만기롤
DERIVED_FIELDS = ["fx_cum_ktb3", "fx_cum_ktb10", "days_to_expiry", "roll_flag"]

# ── 소스별 담당 필드 (실측으로 확정된 최종 배정) ─────────────────────
# 선물 전부(시세·OHLC·거래량·OI·베이시스·수급) → 인포맥스 국채연결선물 C65/C67 (2000/2008+ 연속).
# 국고채금리·CD·KOFR·환율 → ECOS(2000+), 미국채 → FRED, CDS → 수기(라이선스).
SOURCE_FIELDS = {
    "infomax": [   # 국채연결선물 C65/C67 + 스왑
        "ktb3_settle", "ktb3_chg", "ktb10_settle", "ktb10_chg",
        "ktb3_vol", "ktb10_vol", "ktb3_oi", "ktb10_oi",
        "ktb3_open", "ktb3_high", "ktb3_low",
        "ktb10_open", "ktb10_high", "ktb10_low",
        "ktb3_theo_basis", "ktb10_theo_basis",
        "fx_net_ktb3", "fx_net_ktb10",
        "bank_net_ktb3", "itrust_net_ktb3", "ins_net_ktb3",
        "bank_net_ktb10", "itrust_net_ktb10", "ins_net_ktb10",
        "irs3y", "irs10y", "ois3y", "kofr",
        "cpi_yoy", "core_cpi_yoy", "bei", "exports_yoy", "ip_yoy",   # MACRO(ECO)
        "ust10",   # 미국채 10년 지표금리(IR/US10Y MID_Close). FRED(DGS10) 대체 — 지연 없음. as-of 시프트.
        "wti",     # WTI 유가 종가(FRN/SPT:CL 현재가). 미국일→다음 한국거래일 as-of 시프트.
        # v2.6 수급 세분화 (FUT 투자자별 순매수수량)
        "sec_net_ktb3", "sec_net_ktb10",
        "indiv_net_ktb3", "indiv_net_ktb10",
        "etc_net_ktb3", "etc_net_ktb10",
        "pension_net_ktb3", "pension_net_ktb10",
        # v2.7 미국채 OHLC (IR/US10Y·US02Y MID_*) — as-of 시프트
        "ust10_open", "ust10_high", "ust10_low",
        "ust2", "ust2_open", "ust2_high", "ust2_low",
        # v2.8 USDKRW 정규장 OHLC (FX/USDSP_SMBCC 시/고/저/현재가) — 한국날짜(시프트 없음)
        "usdkrw_open", "usdkrw_high", "usdkrw_low", "usdkrw_close",
    ],
    "ecos": ["base_rate", "cd91", "y_ktb3", "y_ktb10", "y_ktb30"],
    "fx":   ["usdkrw"],
    "manual": ["kr_cds5"],   # S&P 라이선스로 자동 불가 → manual_input.csv
}

# ── MACRO 5종 (인포맥스 ECO) — 지수/금액에서 yoy 계산 + 발표일 배치 ──────
# (필드: (종목코드, kind, place))
#  kind : 'yoy'=현재가(지수/금액)에서 전년동월비 계산 / 'level'=현재가 그대로
#  place: 참조월 M 기준 '발표일' 배치 규칙(그 날짜 이후 첫 거래일부터 ffill).
#    'd5'=익월 5일 / 'm1'=익월 1일 / 'm2'=익익월 1일 / 'd16'=익월 16일
#  ※ 차장님 v2.5 실측 역산(cpi·core·bei·exports 100% 바이트 일치, ip는 개정차만).
MACRO_SPEC = {
    "cpi_yoy":      ("701979", "yoy",   "d5"),   # CPI 총지수
    "core_cpi_yoy": ("701996", "yoy",   "d5"),   # 근원 CPI 지수
    "bei":          ("703923", "level", "m1"),   # 기대인플레율(서베이)
    "exports_yoy":  ("557150", "yoy",   "d16"),  # 수출액(금액)
    "ip_yoy":       ("704507", "level", "m2"),   # 전산업생산 yoy
}
MACRO_FIELDS = list(MACRO_SPEC)

# 거래일 앵커 = KTB 선물이 실제 거래된 날. 이 중 하나라도 있는 날만 행 생성(export).
# (base_rate는 캘린더-매일, ust10은 미국캘린더, days_to_expiry/roll_flag는 매일 계산 →
#  이런 건 거래일을 만들지 못하게. KTB 선물 존재 = 한국 국채선물 거래일.)
TRADING_ANCHOR_FIELDS = {
    "ktb3_settle", "ktb3_chg", "ktb3_vol", "ktb3_oi",
    "ktb3_open", "ktb3_high", "ktb3_low",
    "ktb10_settle", "ktb10_chg", "ktb10_vol", "ktb10_oi",
    "ktb10_open", "ktb10_high", "ktb10_low",
}

# 모든 수집필드가 정확히 한 소스에 매핑되는지 검증(파생 제외)
_owned = [f for fs in SOURCE_FIELDS.values() for f in fs]
_expected = [f for f in DATA_FIELDS if f not in DERIVED_FIELDS]
assert sorted(_owned) == sorted(_expected), (
    f"소스 매핑 누락/중복: {set(_expected) ^ set(_owned)}")

# ── 인트라데이(5분봉) 사양 (단일 진실) ────────────────────────────────
# 저장: db/intraday.sqlite, wide bars 테이블 (INTRADAY_DB_PATH). 증분 전용.
# 선물=가격 / 현물=수익률(%). 검증완료 IMDH 레시피는 sandbox/ 참고.
INTRADAY_CYCLE = 5                         # 분봉 주기 (5분)
INTRADAY_SESSION = ("09:00:00", "15:45:00")  # 파이썬 날짜/시각 필터 경계 (B1/D1 상한 안 잘림)
# ★실측: IMDH 분봉은 B1/D1(날짜창)이 무시되고 오직 count개가 '오늘 기준 뒤로' 온다.
#  → 날짜 페이지네이션 불가. count가 유일한 레버. 접근 가능 이력 = 최근 count봉.
#  선물 99999봉 ≈ 2021-07~(5년) / 현물 ≈ 2023-06~. 한 콜 15~17초.
INTRADAY_BACKFILL_START = "2018-01-01"     # 저장 하한(가용이 더 얕으면 그대로). floor 필터용.
INTRADAY_BACKFILL_COUNT = 99999            # 백필: 최대콜 = 가용 전이력 (IMDH 상한)
INTRADAY_INCR_COUNT = 3000                 # 증분: 최근 N봉만(공백 크면 자동 99999로 승격)

# symbol → 조회 스펙. items 순서 = bars 컬럼 매핑 순서(store가 참조).
INTRADAY_SYMBOLS = {
    # 국채 연결선물 (코드 안 바뀜). OI가 분봉으로도 옴.
    "C65": {"kind": "FUT", "label": "KTB3 3년선물",
            "items": ["일자", "시간", "시가", "고가", "저가", "현재가", "거래량", "미결제약정수량"],
            "per": "MM", "real": "false"},
    "C67": {"kind": "FUT", "label": "KTB10 10년선물",
            "items": ["일자", "시간", "시가", "고가", "저가", "현재가", "거래량", "미결제약정수량"],
            "per": "MM", "real": "false"},
    # 현물 지표연속 (2/3/5/10년). 수익률·체결거래량. OI·체결거래대금 없음(미산출).
    "KR1035G00002": {"kind": "BND", "label": "국고2년 지표",
                     "items": ["일자", "시간", "시가", "고가", "저가", "현재가", "체결거래량"],
                     "per": "분", "real": "false"},
    "KR1035G00003": {"kind": "BND", "label": "국고3년 지표",
                     "items": ["일자", "시간", "시가", "고가", "저가", "현재가", "체결거래량"],
                     "per": "분", "real": "false"},
    "KR1035G00005": {"kind": "BND", "label": "국고5년 지표",
                     "items": ["일자", "시간", "시가", "고가", "저가", "현재가", "체결거래량"],
                     "per": "분", "real": "false"},
    "KR1035G00010": {"kind": "BND", "label": "국고10년 지표",
                     "items": ["일자", "시간", "시가", "고가", "저가", "현재가", "체결거래량"],
                     "per": "분", "real": "false"},
}

# ── 수집 기간 / 기준일 ───────────────────────────────────────────────
# 최초 백필 시작일. 소스별로 그 소스 최초제공일이 더 늦으면 자동으로 그때부터(빈 응답 스킵).
#  - ECOS(금리·환율): 2000+ 제공  - KRX(선물): ~2011부터  - 10년선물 상장: 2008-02
from datetime import date as _date
BACKFILL_START = _date(2000, 1, 1)   # 가용 최대(3년선물·국고금리 2000+, 10년선물 2008+)
# 당일 반영 게이트: 이 시각 이후 실행하면 '당일'을 기준일로 인정 → 당일 선물 종가부터 선반영.
# 선물 마감 15:45 직후(16시)로 낮춤 → 마감 나오면 즉시 당일 채우고, 늦는 필드(ECOS 등)는
# 이후 실행에서 idempotent 보강(progressive fill). 15:45 이전엔 당일 미포함(장중 현재가≠종가).
CLOSE_HOUR = 16
# 일별 새로고침 조회건수. IMDH는 count 기반(sort=D=최신순)이라 최근 N거래일만 재조회 = 증분.
# 9000(≈전체 26년) 재계산이 아침 COM busy(RPC_E_CALL_REJECTED) 원인이었음 → 최근분만.
# 월간 MACRO 시트엔 N개월(=10년) → yoy 계산 충분. env IMX_DAILY_COUNT로 오버라이드(전체 재백필 시 크게).
DAILY_REFRESH_COUNT = 120

# HTTP 소스(ecos·fx) 재조회 룩백 = 최근 N거래일은 증분창과 무관하게 항상 다시 긁는다.
# ★왜: last_date_for_source는 그 소스 '모든 필드의 MAX'다. 한 소스 안에서도 공표 시각이
#   다르면(실측: ECOS는 cd91이 국고금리보다 먼저 뜬다) 먼저 온 필드가 창을 닫아버려
#   늦게 나온 필드를 영영 안 긁는다 — 조용한 영구 구멍. (2026-08-12 y_ktb3/10/30 실제 발생)
#   최근 N일 재조회로 늦공표·사후정정을 매 실행마다 흡수한다. upsert가 멱등이라 부작용 없음.
#   비용: ECOS 필드당 1콜에 며칠치가 같이 오므로 사실상 추가 비용 없음.
RECHECK_TRADING_DAYS = 3


_XKRX = None


def _xkrx():
    """한국거래소 공식 캘린더(XKRX) 싱글턴. exchange_calendars 사용."""
    global _XKRX
    if _XKRX is None:
        import exchange_calendars as xcals
        _XKRX = xcals.get_calendar("XKRX")
    return _XKRX


def _weekdays(start, end):
    """[start, end] 평일(월~금). 캘린더가 커버 못 하는 구간 보충용 — 공휴일 미반영."""
    out, d = [], start
    while d <= end:
        if d.weekday() < 5:
            out.append(d)
        d = d + timedelta(days=1)
    return out


def trading_days(start, end):
    """[start, end] 실제 거래일(공휴일 반영). 캘린더 범위 밖 구간만 평일로 보충.

    ★버그 이력(2026-08-18 수정): XKRX는 2006-08-18부터만 제공한다. 예전 구현은
      sessions_in_range(start, end)를 그대로 호출해서, BACKFILL_START(2000-01-01)로
      부르면 DateOutOfBounds가 나고 except가 이를 삼켜 **전 구간이 평일 폴백**이 됐다.
      → 한국 휴장일 286일(2006~2026 실측)을 거래일로 오인. _read_asof(ust10·ust2·wti)와
        _read_macro가 이 함수를 쓰므로 미국 종가가 한국 휴장일에 갇혀 유실되기도 했다
        (실측 피해 1일: 2026-05-26, 한국 대체공휴일 = 미국 Memorial Day가 겹친 날).
      이제 캘린더 하한/상한으로 클램프하고 그 바깥만 평일로 채운다.
    """
    try:
        cal = _xkrx()
        lo, hi = cal.first_session.date(), cal.last_session.date()
        s, e = max(start, lo), min(end, hi)
        sess = [x.date() for x in cal.sessions_in_range(str(s), str(e))] if s <= e else []
        head = _weekdays(start, min(end, lo - timedelta(days=1))) if start < lo else []
        tail = _weekdays(max(start, hi + timedelta(days=1)), end) if end > hi else []
        return head + sess + tail
    except Exception:          # exchange_calendars 미설치 등 → 전 구간 평일
        return _weekdays(start, end)


def ref_date(now=None):
    """기준일 = 마지막으로 '완료된' 거래일 (XKRX 공휴일 반영).
    - 오늘이 거래일이고 CLOSE_HOUR 이후 → 당일 (밤 배치)
    - 그 전(아침)·주말·공휴일 → 직전 거래일
    아침에 돌려도 장 안 열린 당일/휴일을 안 긁게 한다."""
    now = now or datetime.now()
    today = now.date()
    try:
        sess = trading_days(today - timedelta(days=20), today)
        if not sess:
            raise ValueError("no sessions")
        last = sess[-1]                       # today 이하 마지막 거래일
        if last == today and now.hour < CLOSE_HOUR:
            return sess[-2] if len(sess) >= 2 else last
        return last
    except Exception:                          # 폴백: 평일 기준
        d = today
        if not (d.weekday() < 5 and now.hour >= CLOSE_HOUR):
            d = d - timedelta(days=1)
        while d.weekday() >= 5:
            d = d - timedelta(days=1)
        return d


# ── 완결성/무결성 점검용 ──────────────────────────────────────────────
# 소스별 '대표 필드' (이게 있으면 그 소스가 그날 값을 준 것). fred(미국캘린더)는 제외.
KEY_FIELD = {"선물": "ktb3_settle", "금리": "cd91", "환율": "usdkrw", "스왑": "irs3y"}

# ── 수급 항등식 (차장님 요청 — 적재 정합성 자동검증) ────────────────────
# 국채선물은 제로섬이라 전 투자자 순매수 합 = 0. 인포맥스 분류를 우리 컬럼에 사상하면:
#   fx(외국인합) + bank + itrust + ins + sec(증권/선물) + indiv(개인) + etc(나머지 전부) = 0
# ★pension은 etc의 부분집합이므로 여기 넣지 않는다(넣으면 이중계산으로 절대 안 맞음).
# 실측: KTB10 전 구간(3,889일) 불일치 0 / KTB3는 2004-06-28 이후 5,466일 불일치 0.
# 그 이전(2001-01-30~2004-06-25)은 인포맥스가 연기금·종신금·국가/지방을 별도 제공하지 않아
# 원천에서부터 합이 안 맞는다 → 검증 대상에서 제외(데이터는 정상 적재).
FLOW_IDENTITY = {
    "ktb3": ["fx_net_ktb3", "bank_net_ktb3", "itrust_net_ktb3", "ins_net_ktb3",
             "sec_net_ktb3", "indiv_net_ktb3", "etc_net_ktb3"],
    "ktb10": ["fx_net_ktb10", "bank_net_ktb10", "itrust_net_ktb10", "ins_net_ktb10",
              "sec_net_ktb10", "indiv_net_ktb10", "etc_net_ktb10"],
}
FLOW_IDENTITY_START = "2004-06-28"   # 이 날부터 검증 (그 이전은 원천 미분류)

# 0 처리는 인포맥스 로드단에서: 가격필드(settle/OHLC/basis) 0 = 미기록 → 빈칸,
# vol/oi/net 0 = 진짜(거래없음/순매수0) → 유지. sanity 격리는 안 씀.
ZERO_IMPLAUSIBLE = set()

# ── Sanity check 임계값 (사양서 §5) ──────────────────────────────────
# 전일 대비 과대변화 상한 (|Δ|). 초과 시 **플래그(적재+경고)** — 드롭 아님.
# 그로스 오류(자릿수 실수 등)만 잡을 크기로. 실제 급변·롤갭은 플래그로 살려둠.
RANGE_LIMITS = {
    "ktb3_settle": 1.5, "ktb10_settle": 3.0,          # 롤갭 고려 (연결선물)
    "ktb3_open": 1.5, "ktb3_high": 1.5, "ktb3_low": 1.5,
    "ktb10_open": 3.0, "ktb10_high": 3.0, "ktb10_low": 3.0,
    "y_ktb3": 1.0, "y_ktb10": 1.0, "y_ktb30": 1.0,     # 금리 100bp 초과만
    "base_rate": 1.0, "cd91": 1.0, "kofr": 1.0,
    "irs3y": 1.0, "irs10y": 1.0, "ois3y": 1.0, "ust10": 1.0,
    "ust10_open": 1.0, "ust10_high": 1.0, "ust10_low": 1.0,
    "ust2": 1.0, "ust2_open": 1.0, "ust2_high": 1.0, "ust2_low": 1.0,
    "wti": 15.0,    # 유가 $/배럴 일변동 상한(자릿수 오류만 플래그; 실제 급변·2020폭락은 살림)
    # usdkrw는 비율로 따로 처리
}
USDKRW_PCT_LIMIT = 0.10   # 10% 초과만 (2008 위기 등 실제 급변은 살림)

# 동결 감지 대상 = '라이브 피드라 매 거래일 반드시 움직여야 하는' 것만.
# 공표 EOD 지표금리(yield/ust10/usdkrw)는 소수 반올림으로 전일과 우연히 같을 수 있어
# 정상 → 제외 (실측 결과 2년간 51건 전부 오탐이었음).
# 동결은 '드롭'이 아니라 '플래그+적재'(사양서 §5: 의심→플래그). sanity.py 참조.
FREEZE_CHECK_FIELDS = {
    "ktb3_settle", "ktb10_settle",
}

# 음수 불가 (vol·oi는 0 이상 정수)
NONNEG_INT_FIELDS = {"ktb3_vol", "ktb10_vol", "ktb3_oi", "ktb10_oi"}

# 양수여야 하는 금리류 (>0)
POSITIVE_FIELDS = {
    "y_ktb3", "y_ktb10", "y_ktb30", "base_rate", "cd91", "kofr",
    "irs3y", "irs10y", "ois3y", "ust10", "usdkrw",
    "ust10_open", "ust10_high", "ust10_low",
    "ust2", "ust2_open", "ust2_high", "ust2_low",
    "usdkrw_open", "usdkrw_high", "usdkrw_low", "usdkrw_close",   # FX 환율 >0
}
