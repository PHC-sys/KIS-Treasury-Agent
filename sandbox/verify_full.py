# -*- coding: utf-8 -*-
"""07-31 5분봉(선물3/10 + 현물3/10) → 일별 전 항목 대조."""
import win32com.client as win32
import pythoncom, time, sys, csv, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull
D = next(r for r in csv.DictReader(open(r"D:\KIS Treasury Agent\out\ktb_daily.csv",encoding="utf-8")) if r["date"]=="2026-07-31")
FUT=["일자","시간","시가","고가","저가","현재가","거래량","미결제약정수량"]
CASH=["일자","시간","시가","고가","저가","현재가","체결거래량","체결거래대금"]
pythoncom.CoInitialize()
app=win32.Dispatch("Excel.Application")
for _ in range(15):
    try: app.Version; break
    except: time.sleep(1)
app.DisplayAlerts=False; app.Visible=False
infomax_pull._register_addin(app)
wb=app.Workbooks.Add(); ws=wb.Worksheets(1)
ws.Range("B1").Value=dt.datetime(2026,7,31,9,0); ws.Range("D1").Value=dt.datetime(2026,7,31,15,45); ws.Range("F1").Value=99999
specs=[("FUT","C65",FUT,"MM",1),("FUT","C67",FUT,"MM",11),("BND","KR1035G00003",CASH,"분",21),("BND","KR1035G00010",CASH,"분",31)]
for mkt,sym,items,per,c0 in specs:
    for j,it in enumerate(items): ws.Cells(3,c0+j).Value=it
    a=chr(ord('A')+c0-1); b=chr(ord('A')+c0-1+len(items)-1)
    opt=f"Per={per},Cycle=5,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
    ws.Cells(2,c0).Formula=f'=IMDH("{mkt}","{sym}",{a}3:{b}3,$B$1,$D$1,$F$1,"{opt}")'
time.sleep(16)
def blk(c0):
    d=ws.Range(ws.Cells(4,c0),ws.Cells(700,c0+7)).Value
    return [r for r in d if r[0] is not None and str(r[0])[:10]=="2026-07-31"]
def fnum(x):
    try: return float(x)
    except: return None
def chk(name, got, exp, tol=0.001):
    e=float(exp); ok=got is not None and abs(got-e)<=tol
    print(f"  {name:16} 5분={got}  일별={exp}  {'OK' if ok else 'X'}")
f3,f10,c3,c10=blk(1),blk(11),blk(21),blk(31)
print(f"07-31 봉수: 선물3={len(f3)} 선물10={len(f10)} 현물3={len(c3)} 현물10={len(c10)}")
print("[선물3년]"); chk("종가",fnum(f3[0][5]),D["ktb3_settle"]); chk("고",max(fnum(r[3]) for r in f3),D["ktb3_high"]); chk("저",min(fnum(r[4]) for r in f3),D["ktb3_low"]); chk("거래량합",sum(fnum(r[6]) or 0 for r in f3),D["ktb3_vol"],tol=1); chk("OI",fnum(f3[0][7]),D["ktb3_oi"],tol=1)
print("[선물10년]"); chk("종가",fnum(f10[0][5]),D["ktb10_settle"]); chk("고",max(fnum(r[3]) for r in f10),D["ktb10_high"]); chk("저",min(fnum(r[4]) for r in f10),D["ktb10_low"]); chk("거래량합",sum(fnum(r[6]) or 0 for r in f10),D["ktb10_vol"],tol=1); chk("OI",fnum(f10[0][7]),D["ktb10_oi"],tol=1)
print("[현물3년 수익률]"); chk("종가수익률",fnum(c3[0][5]),D["y_ktb3"],tol=0.02)
print("[현물10년 수익률]"); chk("종가수익률",fnum(c10[0][5]),D["y_ktb10"],tol=0.02)
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
