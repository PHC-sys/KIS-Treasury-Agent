# -*- coding: utf-8 -*-
"""현물 국고채(장외) 5분 되는지. 구분=BND, 지표연속 코드, 장외 수익률 OHLC 추정 항목."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull

def diag(mkt, sym, items, per, label):
    print("="*60); print(f"{label} | 구분={mkt} 종목={sym} Per={per}")
    print("항목:", items)
    pythoncom.CoInitialize()
    app = win32.Dispatch("Excel.Application")
    for _ in range(15):
        try: app.Version; break
        except: time.sleep(1)
    app.DisplayAlerts=False; app.Visible=False
    infomax_pull._register_addin(app)
    wb = app.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Range("B1").Value = dt.datetime(2026,8,1)
    ws.Range("D1").Value = dt.datetime.combine(dt.date.today(), dt.time())
    ws.Range("F1").Value = 99999
    for j, it in enumerate(items): ws.Cells(3, j+1).Value = it
    last = chr(ord("A")+len(items)-1)
    cyc = ",Cycle=5" if per=="MM" else ""
    opt = f"Per={per}{cyc},sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
    ws.Range("A2").Formula = f'=IMDH("{mkt}","{sym}",A3:{last}3,$B$1,$D$1,$F$1,"{opt}")'
    time.sleep(12)
    print("A2(상태):", ws.Range("A2").Value)
    for r in range(4, 10):
        print("  ", [str(ws.Cells(r,c).Value)[:14] if ws.Cells(r,c).Value is not None else '' for c in range(1,len(items)+1)])
    try: wb.Close(SaveChanges=False)
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    diag("BND", "KR1035G00003",
         ["일자","시간","장외-시 수익률","장외-고 수익률","장외-저 수익률","장외-현재수익률","장외-누적거래량"],
         "MM", "국고3년 장외 5분(수익률 추정)")
