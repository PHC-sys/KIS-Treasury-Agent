# -*- coding: utf-8 -*-
"""장외 항목이 일별로는 값이 오나 + 5분에선 뭐가 오나 비교."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull

ITEMS = ["일자","시간","장외-시 수익률","장외-고 수익률","장외-저 수익률","장외-종 수익률","장외-누적거래량"]

def run(app, ws, per, label):
    for j, it in enumerate(ITEMS): ws.Cells(3, j+1).Value = it
    last = chr(ord("A")+len(ITEMS)-1)
    cyc = ",Cycle=5" if per=="MM" else ""
    opt = f"Per={per}{cyc},sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
    ws.Range("A2").Formula = f'=IMDH("BND","KR1035G00003",A3:{last}3,$B$1,$D$1,$F$1,"{opt}")'
    time.sleep(12)
    print(f"--- {label} (Per={per}) --- A2:", ws.Range("A2").Value)
    for r in range(4, 10):
        print("  ", [str(ws.Cells(r,c).Value)[:13] if ws.Cells(r,c).Value is not None else '' for c in range(1,len(ITEMS)+1)])

pythoncom.CoInitialize()
app = win32.Dispatch("Excel.Application")
for _ in range(15):
    try: app.Version; break
    except: time.sleep(1)
app.DisplayAlerts=False; app.Visible=False
infomax_pull._register_addin(app)
wb = app.Workbooks.Add()
ws = wb.Worksheets(1)
ws.Range("B1").Value = dt.datetime(2026,7,20)
ws.Range("D1").Value = dt.datetime.combine(dt.date.today(), dt.time())
ws.Range("F1").Value = 99999
run(app, ws, "D", "일별")
ws2 = wb.Worksheets.Add()
ws2.Range("B1").Value = dt.datetime(2026,7,31)
ws2.Range("D1").Value = dt.datetime(2026,7,31,23,59)
ws2.Range("F1").Value = 99999
run(app, ws2, "MM", "7/31 하루 5분")
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
