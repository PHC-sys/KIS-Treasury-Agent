# -*- coding: utf-8 -*-
"""07-31 5분봉(선물 C65 + 현물 KR1035G00003)으로 그날 OHLC/OI/수익률 산출 → 일별과 대조."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull

def pull(ws, col0, mkt, sym, items, per):
    for j,it in enumerate(items): ws.Cells(3, col0+j).Value=it
    c0=chr(ord("A")+col0-1); c1=chr(ord("A")+col0-1+len(items)-1)
    opt=f"Per={per},Cycle=5,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
    ws.Cells(2,col0).Formula=f'=IMDH("{mkt}","{sym}",{c0}3:{c1}3,$B$1,$D$1,$F$1,"{opt}")'

pythoncom.CoInitialize()
app=win32.Dispatch("Excel.Application")
for _ in range(15):
    try: app.Version; break
    except: time.sleep(1)
app.DisplayAlerts=False; app.Visible=False
infomax_pull._register_addin(app)
wb=app.Workbooks.Add(); ws=wb.Worksheets(1)
ws.Range("B1").Value=dt.datetime(2026,7,31)
ws.Range("D1").Value=dt.datetime(2026,7,31,23,59)
ws.Range("F1").Value=99999
FUT=["일자","시간","시가","고가","저가","현재가","거래량","미결제약정수량"]
CASH=["일자","시간","시가","고가","저가","현재가","체결거래량","체결거래대금"]
pull(ws, 1,  "FUT", "C65", FUT, "MM")            # A~H
pull(ws, 11, "BND", "KR1035G00003", CASH, "분")  # K~R
time.sleep(15)
def read(col0, n):
    d=ws.Range(ws.Cells(4,col0), ws.Cells(500, col0+n-1)).Value
    return [r for r in d if r[0] is not None]
fut=read(1,8); cash=read(11,8)
def f(x):
    try: return float(x)
    except: return None
# 선물 그날: 종가=최신봉 현재가, 고=max(고가), 저=min(저가), 거래량합, OI=최신봉
fh=[f(r[3]) for r in fut if f(r[3])]; fl=[f(r[4]) for r in fut if f(r[4])]
print("=== 선물 C65 (07-31) 5분→일 ===")
print(f"  종가(최신봉 현재가)={fut[0][5]}  vs 일별 ktb3_settle=103.23")
print(f"  고={max(fh)} vs 103.25 | 저={min(fl)} vs 103.06")
print(f"  거래량합={sum(f(r[6]) or 0 for r in fut):.0f} vs 167253 | OI(최신봉)={fut[0][7]} vs 597425")
print(f"  봉 수={len(fut)}")
ch=[f(r[3]) for r in cash if f(r[3])]; cl=[f(r[4]) for r in cash if f(r[4])]
print("=== 현물 KR1035G00003 국고3년 (07-31) ===")
print(f"  수익률 종가(최신봉)={cash[0][5]}  vs 일별 y_ktb3=3.758")
print(f"  고={max(ch)} 저={min(cl)} | 봉 수={len(cash)}")
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
