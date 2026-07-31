# -*- coding: utf-8 -*-
"""통합 인포맥스 워크북(infomax_data.xlsx) 생성 — 필요한 항목만 시트별 IMDH.
종목·항목 정의의 단일 진실. 종목 추가/변경 시 SHEETS 고치고 재실행."""
import win32com.client as win32
import pythoncom
import time
import os
import datetime as dt
import infomax_pull
import config

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "infomax_data.xlsx")

# 국채연결선물(C65=3년, C67=10년) — 정산가·OHLC·거래량·OI·베이시스·수급 전부 2000/2008+ 연속.
_FUT = ["일자", "현재가", "전일대비", "시가", "고가", "저가", "거래량",
        "미결제약정수량", "이론베이시스",
        "외국인합매수수량", "외국인합매도수량", "은행매수수량", "은행매도수량",
        "투신매수수량", "투신매도수량", "보험매수수량", "보험매도수량"]

# (시트명, 시장구분, 종목, [항목들; 첫 항목은 일자])
SHEETS = [
    ("FUT3",   "FUT", "C65", _FUT),
    ("FUT10",  "FUT", "C67", _FUT),
    ("IRS3Y",  "IR",  "IRSTP1KRW03Y", ["일자", "MID종가"]),
    ("IRS10Y", "IR",  "IRSTP1KRW10Y", ["일자", "MID종가"]),
    ("OIS3Y",  "IR",  "IRSKM5KRW03Y", ["일자", "MID종가"]),
]


def col_letter(n):
    return chr(ord("A") + n - 1)


def main():
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


if __name__ == "__main__":
    main()
