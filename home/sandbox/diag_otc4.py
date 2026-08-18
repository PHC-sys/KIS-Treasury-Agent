# -*- coding: utf-8 -*-
"""국고10년 장외 5분 + Cyclesync=true(거래없는 봉 이전값 carry). COM 직접 읽기."""
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
ws.Range("B1").Value = dt.datetime(2026,8,1)
ws.Range("D1").Value = dt.datetime.combine(dt.date.today(), dt.time(23,59))
ws.Range("F1").Value = 99999
for j,it in enumerate(ITEMS): ws.Cells(3,j+1).Value=it
opt = "Per=MM,Cycle=5,Cyclesync=true,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
ws.Range("A2").Formula = f'=IMDH("BND","KR1035G00010",A3:G3,$B$1,$D$1,$F$1,"{opt}")'
time.sleep(14)
print("A2:", ws.Range("A2").Value)
nz=0; tot=0
for r in range(4, 24):
    v = ws.Cells(r,6).Value   # 종수익률
    d = ws.Cells(r,1).Value; t = ws.Cells(r,2).Value
    if d is None: break
    tot+=1
    try:
        if float(v)>0: nz+=1
    except: pass
    if r<14:
        hh=int(float(t or 0)*24); mm=int((float(t or 0)*24-hh)*60) if t else 0
        print(f"  {str(d)[:10]} {hh:02d}:{mm:02d}  시={ws.Cells(r,3).Value} 고={ws.Cells(r,4).Value} 저={ws.Cells(r,5).Value} 종={v} 누적={ws.Cells(r,7).Value}")
print(f"상위 {tot}봉 중 종≠0: {nz}")
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
