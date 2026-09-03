# -*- coding: utf-8 -*-
"""인포맥스 워크북(infomax_data.xlsx) — 시트별 IMDH. 종목 정의의 단일 진실.

- `python build_infomax.py`        : 워크북 통째 재생성(전 시트, 무거움)
- `python build_infomax.py add`    : 기존 워크북에 없는 시트만 추가(신규 시트만 계산, 가벼움)
종목 추가/변경 시 SHEETS 고치고 `add`로 실행하면 됨.
"""
import win32com.client as win32
import pythoncom
import time
import os
import sys
import datetime as dt
import infomax_pull
import config

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "infomax_data.xlsx")

# 국채연결선물(C65=3년, C67=10년) — 정산가·OHLC·거래량·OI·베이시스·수급 전부 2000/2008+ 연속.
_FUT = ["일자", "현재가", "전일대비", "시가", "고가", "저가", "거래량",
        "미결제약정수량", "이론베이시스",
        "외국인합매수수량", "외국인합매도수량", "은행매수수량", "은행매도수량",
        "투신매수수량", "투신매도수량", "보험매수수량", "보험매도수량",
        # v2.6 수급 세분화 — 인포맥스가 '순매수수량'을 직접 주므로 매수/매도 차분 불필요.
        # (기존 4종과 값이 동일함은 전 구간 대조로 확인: KTB3 6,305일·KTB10 3,889일 불일치 0)
        # etc_net = 기타+국가/지방+연기금+종신금 이라 네 항목을 모두 받는다.
        "증권/선물순매수수량", "개인순매수수량",
        "연기금순매수수량", "기타순매수수량", "국가/지방순매수수량", "종신금순매수수량"]

# (시트명, 시장구분, 종목, [항목들; 첫 항목은 일자])
SHEETS = [
    ("FUT3",   "FUT", "C65", _FUT),
    ("FUT10",  "FUT", "C67", _FUT),
    ("IRS3Y",  "IR",  "IRSTP1KRW03Y", ["일자", "MID종가"]),
    ("IRS10Y", "IR",  "IRSTP1KRW10Y", ["일자", "MID종가"]),
    ("OIS3Y",  "IR",  "IRSKM5KRW03Y", ["일자", "MID종가"]),
    ("KOFR",   "ECO", "785831",       ["일자", "현재가"]),   # 한국 KOFR금리 (ECOS보다 신선)
    # 미국채 10년 지표금리(on-the-run) MID종가 → ust10. FRED(DGS10) 대체: 지연 없이 당일 확보.
    # 미국일자로 옴 → infomax_pull._read_ust10이 '다음 한국거래일'로 as-of 시프트(룩어헤드 방지).
    # ★만기는 2자리 제로패딩: US02Y·US10Y·US20Y·US30Y (US2Y/US5Y 같은 1자리는 빈 응답).
    #  단말 화면 코드 GVO:TR10Y/GVO:TR02Y는 시장구분이 FRN이고 '현재가'=BID_Close와 동일.
    #  우리는 IR 경로로 MID를 쓴다(기존 ust10 26년치가 MID_Close라 시리즈 통일).
    ("US10Y",  "IR",  "US10Y",  ["일자", "MID_Open", "MID_High", "MID_Low", "MID_Close"]),
    ("US02Y",  "IR",  "US02Y",  ["일자", "MID_Open", "MID_High", "MID_Low", "MID_Close"]),
    # WTI 유가 종가(현재가) → wti. 미국일자 → _read_wti가 다음 한국거래일로 as-of 시프트.
    ("WTI",    "FRN", "SPT:CL",         ["일자", "현재가"]),
    # USDKRW 정규장(09:00~15:30) OHLC → usdkrw_open/high/low/close. 한국날짜(시프트 없음).
    # ★_EXT(연장/야간) 아니라 정규 USDSP_SMBCC. close=현재가(15:30 마감). LOAD_SPEC이 매핑.
    ("USDKRW", "FX",  "USDSP_SMBCC",    ["일자", "시가", "고가", "저가", "현재가"]),
    # ── MACRO 5종 (ECO, 월별) — infomax_pull이 지수→yoy 계산 + 발표일 배치 ──
    ("M_CPI",  "ECO", "701979",       ["일자", "현재가"]),   # CPI 총지수 → cpi_yoy
    ("M_CORE", "ECO", "701996",       ["일자", "현재가"]),   # 근원 CPI 지수 → core_cpi_yoy
    ("M_BEI",  "ECO", "703923",       ["일자", "현재가"]),   # 기대인플레율 → bei
    ("M_EXP",  "ECO", "557150",       ["일자", "현재가"]),   # 수출액 → exports_yoy
    ("M_IP",   "ECO", "704507",       ["일자", "현재가"]),   # 전산업생산 → ip_yoy
]


def col_letter(n):
    return chr(ord("A") + n - 1)


def _write_sheet(ws, name, mkt, sym, items, start, ref):
    """시트 하나에 파라미터(1행)·항목헤더(3행)·IMDH 수식(A2) 기입."""
    ws.Name = name
    for cell, val in [("A1", "시작"), ("B1", start), ("C1", "종료"), ("D1", ref),
                      ("E1", "Data 개수"), ("F1", 9000), ("G1", "주기"), ("H1", "일"),
                      ("I1", "정렬"), ("J1", "D"), ("K1", "영업일"), ("L1", 0),
                      ("M1", "시세산출"), ("N1", "종가")]:
        ws.Range(cell).Value = val
    for j, it in enumerate(items):
        ws.Cells(3, j + 1).Value = it
    rng = f"A3:{col_letter(len(items))}3"
    ws.Range("A2").Formula = (
        f'=IMDH("{mkt}","{sym}",{rng},$B$1,$D$1,$F$1,'
        f'"Per="&$H$1&",sort="&$J$1&",real=false,Bizday="&$L$1&'
        f'",Quote="&$N$1&",Pos=20,Orient=V,Title={name},DtFmt=1,TmFmt=1,unit=true")')


def main():
    """워크북 통째 재생성(전 시트 재수집 — 무거움)."""
    pythoncom.CoInitialize()
    app = win32.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    infomax_pull._register_addin(app)
    ref = dt.datetime.combine(config.ref_date(), dt.time())
    start = dt.datetime(2000, 1, 1)        # 가용한 최대 이력 (3년선물 2000+, 10년 2008+)
    wb = app.Workbooks.Add()
    try:
        for i, (name, mkt, sym, items) in enumerate(SHEETS):
            ws = wb.Worksheets(i + 1) if i < wb.Worksheets.Count else \
                wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            _write_sheet(ws, name, mkt, sym, items, start, ref)
        while wb.Worksheets.Count > len(SHEETS):
            wb.Worksheets(wb.Worksheets.Count).Delete()
        app.CalculateFullRebuild(); time.sleep(30)
        app.CalculateFullRebuild(); time.sleep(3)
        if os.path.exists(OUT):
            os.remove(OUT)
        wb.SaveAs(OUT, FileFormat=51)
        print("생성:", OUT)
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


def add_missing():
    """기존 워크북을 열어 SHEETS 중 '없는 시트만' 추가 → 신규 시트만 계산 → 저장.
    기존 26년치는 건드리지 않아 가볍다. (사용자 Excel이 떠 있으면 그 인스턴스 보존)"""
    if not os.path.exists(OUT):
        print(f"  ⚠ {OUT} 없음 — 최초엔 전체 생성 필요(python build_infomax.py)")
        return
    pythoncom.CoInitialize()
    started = False
    try:
        app = win32.GetActiveObject("Excel.Application")   # 떠 있는 Excel에 붙음
    except Exception:
        app = win32.Dispatch("Excel.Application"); started = True
    app.DisplayAlerts = False
    if started:
        app.Visible = False
    infomax_pull._register_addin(app)
    ref = dt.datetime.combine(config.ref_date(), dt.time())
    start = dt.datetime(2000, 1, 1)
    wb = app.Workbooks.Open(OUT)
    try:
        existing = {wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)}
        added = []
        for name, mkt, sym, items in SHEETS:
            if name in existing:
                continue
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            _write_sheet(ws, name, mkt, sym, items, start, ref)
            added.append(name)
        if not added:
            print("추가할 시트 없음 — 이미 다 있음")
            return
        for name in added:                       # 신규 시트만 계산(전체 재수집 회피)
            wb.Worksheets(name).Calculate()
        time.sleep(int(os.environ.get("IMX_WAIT", "30")))
        for name in added:
            wb.Worksheets(name).Calculate()
        time.sleep(5)
        wb.Save()
        print("추가·계산·저장 완료:", added)
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        if started:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "add":
        add_missing()
    else:
        main()
