# -*- coding: utf-8 -*-
"""분봉 진단 v2: 작동하는 일별 옵션(Pos=20,Orient=V 포함) + Per=MM,Cycle=5."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull

def diag(sym, cycle, start, end):
    print("="*60); print(f"{sym} {cycle}분  {start}~{end}")
    pythoncom.CoInitialize()
    app = win32.Dispatch("Excel.Application")
    for _ in range(15):
        try: app.Version; break
        except: time.sleep(1)
    app.DisplayAlerts=False; app.Visible=False
    infomax_pull._register_addin(app)
    wb = app.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Range("B1").Value = dt.datetime.fromisoformat(start)
    ws.Range("D1").Value = dt.datetime.fromisoformat(end)
    ws.Range("F1").Value = 99999
    for j, it in enumerate(["일자","시가","고가","저가","현재가","거래량"]):
        ws.Cells(3, j+1).Value = it
    # 일별 pipeline과 동일 옵션 + Per=MM,Cycle (sort=D: 최근이 위로)
    opt = f"Per=MM,Cycle={cycle},sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
    ws.Range("A2").Formula = f'=IMDH("FUT","{sym}",A3:F3,$B$1,$D$1,$F$1,"{opt}")'
    time.sleep(12)
    print("A2(상태):", ws.Range("A2").Value)
    print("데이터(A4부터, 최근순):")
    for r in range(4, 12):
        vals = [ws.Cells(r, c).Value for c in range(1, 7)]
        print(f"  {[str(v)[:19] if v is not None else '' for v in vals]}")
    try: wb.Close(SaveChanges=False)
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    diag("C65", 5, "2026-08-01", dt.date.today().isoformat())
