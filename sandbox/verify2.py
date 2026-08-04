# -*- coding: utf-8 -*-
"""반환 날짜범위 확인 + 07-31만 필터해서 일별과 대조."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull
pythoncom.CoInitialize()
app=win32.Dispatch("Excel.Application")
for _ in range(15):
    try: app.Version; break
    except: time.sleep(1)
app.DisplayAlerts=False; app.Visible=False
infomax_pull._register_addin(app)
wb=app.Workbooks.Add(); ws=wb.Worksheets(1)
ws.Range("B1").Value=dt.datetime(2026,7,28)
ws.Range("D1").Value=dt.datetime(2026,7,31,23,59)
ws.Range("F1").Value=99999
FUT=["일자","시간","시가","고가","저가","현재가","거래량","미결제약정수량"]
for j,it in enumerate(FUT): ws.Cells(3,j+1).Value=it
opt="Per=MM,Cycle=5,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
ws.Range("A2").Formula=f'=IMDH("FUT","C65",A3:H3,$B$1,$D$1,$F$1,"{opt}")'
time.sleep(14)
data=ws.Range("A4:H600").Value
rows=[r for r in data if r[0] is not None]
def dstr(x): return str(x)[:10]
dates=sorted(set(dstr(r[0]) for r in rows))
print(f"반환 봉수={len(rows)}, 날짜={dates[:3]}...{dates[-3:]} (총 {len(dates)}일)")
d31=[r for r in rows if dstr(r[0])=="2026-07-31"]
print(f"07-31 봉수={len(d31)} (5분이면 ~80)")
def f(x):
    try: return float(x)
    except: return None
if d31:
    # sort=D라 d31[0]=07-31 최신봉(마지막), d31[-1]=첫봉
    hs=[f(r[3]) for r in d31 if f(r[3])]; ls=[f(r[4]) for r in d31 if f(r[4])]
    print(f"  07-31 최신봉 현재가(종가)={d31[0][5]} vs ktb3_settle=103.23")
    print(f"  첫봉(09:00) 시가={d31[-1][2]}")
    print(f"  고={max(hs)} vs 103.25 | 저={min(ls)} vs 103.06")
    print(f"  거래량합={sum(f(r[6]) or 0 for r in d31):.0f} vs 167253 | OI(최신봉)={d31[0][7]} vs 597425")
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
