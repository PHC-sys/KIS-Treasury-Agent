# -*- coding: utf-8 -*-
"""현물 국고3년 5분(Cycle=1) 07-31 → 장외 일별 + y_ktb3(ECOS) 대조."""
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
# 5분(Cycle=1) — 시각 경계
ws.Range("B1").Value=dt.datetime(2026,7,31,9,0); ws.Range("D1").Value=dt.datetime(2026,7,31,15,45); ws.Range("F1").Value=99999
C=["일자","시간","시가","고가","저가","현재가","체결거래량","체결거래대금"]
for j,it in enumerate(C): ws.Cells(3,j+1).Value=it
opt="Per=분,Cycle=1,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
ws.Range("A2").Formula=f'=IMDH("BND","KR1035G00003",A3:H3,$B$1,$D$1,$F$1,"{opt}")'
# 장외 일별 OHLC(같은 소스) 별 시트
ws2=wb.Worksheets.Add()
ws2.Range("B1").Value=dt.datetime(2026,7,25); ws2.Range("D1").Value=dt.datetime(2026,8,4); ws2.Range("F1").Value=99999
DD=["일자","장외-시 수익률","장외-고 수익률","장외-저 수익률","장외-종 수익률"]
for j,it in enumerate(DD): ws2.Cells(3,j+1).Value=it
opt2="Per=일,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,unit=true"
ws2.Range("A2").Formula=f'=IMDH("BND","KR1035G00003",A3:E3,$B$1,$D$1,$F$1,"{opt2}")'
time.sleep(16)
d=ws.Range("A4:H900").Value
c31=[r for r in d if r[0] is not None and str(r[0])[:10]=="2026-07-31"]
def fn(x):
    try: return float(x)
    except: return None
print(f"현물3년 5분(Cycle=1) 07-31 봉수: {len(c31)}")
if c31:
    hs=[fn(r[3]) for r in c31 if fn(r[3])]; ls=[fn(r[4]) for r in c31 if fn(r[4])]
    tr=[r for r in c31 if (fn(r[6]) or 0)>0]  # 체결 있는 봉
    print(f"  종가수익률(최신봉)={c31[0][5]} | 고={max(hs)} 저={min(ls)} | 체결봉수={len(tr)}/{len(c31)}")
# 장외 일별 07-31
dd=ws2.Range("A4:E30").Value
r31=next((r for r in dd if r[0] is not None and str(r[0])[:10]=="2026-07-31"), None)
print(f"장외 일별 07-31: 시={r31[1]} 고={r31[2]} 저={r31[3]} 종={r31[4]}" if r31 else "장외일별 07-31 없음")
print("우리 일별 y_ktb3(ECOS) = 3.758")
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
