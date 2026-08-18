# -*- coding: utf-8 -*-
"""분봉 조회 되는지 + 구조 확인 (sandbox, 본 파이프라인 안 건드림).
FUT C65(KTB3선물) 5분봉을 최근 구간으로 뽑아 sandbox/test_5min.xlsx 저장."""
import win32com.client as win32
import pythoncom, time, os, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull   # _register_addin 재사용

OUT = r"D:\KIS Treasury Agent\sandbox\test_5min.xlsx"

def pull(sym, cycle, start, end, count, items):
    pythoncom.CoInitialize()
    app = win32.Dispatch("Excel.Application")
    for _ in range(15):
        try: app.Version; break
        except: time.sleep(1)
    app.DisplayAlerts = False; app.Visible = False
    blank = app.Workbooks.Add()
    try: app.Calculation = -4135    # 수동
    except: pass
    infomax_pull._register_addin(app)
    wb = app.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Range("B1").Value = dt.datetime.fromisoformat(start)
    ws.Range("D1").Value = dt.datetime.fromisoformat(end)
    ws.Range("F1").Value = count
    for j, it in enumerate(items):
        ws.Cells(3, j+1).Value = it
    last = chr(ord("A")+len(items)-1)
    ws.Range("A2").Formula = (
        f'=IMDH("FUT","{sym}",A3:{last}3,$B$1,$D$1,$F$1,'
        f'"Per=MM,Cycle={cycle},sort=A,real=false,DtFmt=1,TmFmt=1")')
    app.CalculateFullRebuild(); time.sleep(20)
    app.CalculateFullRebuild(); time.sleep(3)
    nrows = ws.UsedRange.Rows.Count
    print("UsedRange 행수:", nrows)
    if os.path.exists(OUT): os.remove(OUT)
    wb.SaveAs(OUT, FileFormat=51)
    print("저장:", OUT)
    try: wb.Close(SaveChanges=False)
    except: pass
    try: blank.Close(SaveChanges=False)
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    # 최근 ~10거래일 5분봉 (구조·개수 확인용)
    pull("C65", 5, "2026-07-21", dt.date.today().isoformat(), 99999,
         ["일자","시가","고가","저가","현재가","거래량"])
