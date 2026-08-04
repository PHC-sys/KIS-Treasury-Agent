# -*- coding: utf-8 -*-
"""국고10년 장외 5분 오늘 하루 통째 읽어 0 아닌 봉 존재/시각 확인."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull
ITEMS = ["일자","시간","장외-시 수익률","장외-고 수익률","장외-저 수익률","장외-종 수익률","장외-누적거래량"]
pythoncom.CoInitialize()
app = win32.Dispatch("Excel.Application")
for _ in range(15):
    try: app.Version; break
    except: time.sleep(1)
app.DisplayAlerts=False; app.Visible=False
infomax_pull._register_addin(app)
wb = app.Workbooks.Add(); ws = wb.Worksheets(1)
ws.Range("B1").Value = dt.datetime(2026,8,4)         # 오늘만
ws.Range("D1").Value = dt.datetime(2026,8,4,23,59)
ws.Range("F1").Value = 99999
for j,it in enumerate(ITEMS): ws.Cells(3,j+1).Value=it
opt = "Per=MM,Cycle=5,sort=A,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
ws.Range("A2").Formula = f'=IMDH("BND","KR1035G00010",A3:G3,$B$1,$D$1,$F$1,"{opt}")'
time.sleep(14)
data = ws.Range("A4:G200").Value    # 한 번에
def num(x):
    try: return float(x)
    except: return None
rows = [r for r in data if r[0] is not None]
nz = [r for r in rows if (num(r[5]) or 0) > 0]
print(f"오늘 5분봉 총 {len(rows)}개, 종수익률≠0: {len(nz)}개")
def hm(t):
    t=num(t) or 0; h=int(t*24); return f"{h:02d}:{int((t*24-h)*60):02d}"
if nz:
    print("0 아닌 봉(시각/시고저종/누적거래량):")
    for r in nz[:6]+nz[-3:]:
        print(f"  {hm(r[1])}  시={r[2]} 고={r[3]} 저={r[4]} 종={r[5]}  누적={r[6]}")
else:
    print("전부 0 — 오늘 장외 10년 5분에 값 없음")
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
