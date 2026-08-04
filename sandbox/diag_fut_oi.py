# -*- coding: utf-8 -*-
"""선물 5분봉에 OI(미결제약정수량)까지 오는지 확인."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull

def diag(sym, items, label):
    print("="*60); print(label, sym, "| 항목:", items)
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
    opt = "Per=MM,Cycle=5,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
    ws.Range("A2").Formula = f'=IMDH("FUT","{sym}",A3:{last}3,$B$1,$D$1,$F$1,"{opt}")'
    time.sleep(12)
    print("A2:", ws.Range("A2").Value, "| 헤더:", [ws.Cells(3,c).Value for c in range(1,len(items)+1)])
    for r in range(4, 9):
        print("  ", [str(ws.Cells(r,c).Value)[:12] if ws.Cells(r,c).Value is not None else '' for c in range(1,len(items)+1)])
    try: wb.Close(SaveChanges=False)
    except: pass
    try: app.Quit()
    except: pass

if __name__ == "__main__":
    diag("C65", ["일자","시간","시가","고가","저가","현재가","거래량","미결제약정수량"], "KTB3선물+OI")
