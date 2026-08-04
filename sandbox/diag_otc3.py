# -*- coding: utf-8 -*-
"""국고10년 장외 5분 — 저장 후 openpyxl로 0 아닌 봉 비율/샘플 분석."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt, openpyxl
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull
OUT = r"D:\KIS Treasury Agent\sandbox\otc_5min.xlsx"
ITEMS = ["일자","시간","장외-시 수익률","장외-고 수익률","장외-저 수익률","장외-종 수익률","장외-누적거래량"]

pythoncom.CoInitialize()
app = win32.Dispatch("Excel.Application")
for _ in range(15):
    try: app.Version; break
    except: time.sleep(1)
app.DisplayAlerts=False; app.Visible=False
infomax_pull._register_addin(app)
wb = app.Workbooks.Add(); ws = wb.Worksheets(1)
ws.Range("B1").Value = dt.datetime(2026,7,28)   # 최근 ~1주 5분
ws.Range("D1").Value = dt.datetime.combine(dt.date.today(), dt.time(23,59))
ws.Range("F1").Value = 99999
for j,it in enumerate(ITEMS): ws.Cells(3,j+1).Value=it
opt = "Per=MM,Cycle=5,sort=A,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
ws.Range("A2").Formula = f'=IMDH("BND","KR1035G00010",A3:G3,$B$1,$D$1,$F$1,"{opt}")'
time.sleep(15)
import os
if os.path.exists(OUT): os.remove(OUT)
wb.SaveAs(OUT, FileFormat=51)
try: wb.Close(SaveChanges=False)
except: pass
try: app.Quit()
except: pass
# 분석
wv = openpyxl.load_workbook(OUT, data_only=True).active
rows=[r for r in wv.iter_rows(min_row=4, values_only=True)]
rows=[r for r in rows if r[0] is not None]
def num(x):
    try: return float(x)
    except: return 0.0
nz = [r for r in rows if num(r[5])>0]   # 종수익률>0
print(f"총 5분봉: {len(rows)}  |  종수익률≠0: {len(nz)}  ({100*len(nz)/max(len(rows),1):.1f}%)")
print("0 아닌 봉 샘플(일자, 시간→시각, 종수익률, 누적거래량):")
for r in nz[-8:]:
    hh = int(num(r[1])*24); mm=int((num(r[1])*24-hh)*60)
    print(f"  {str(r[0])[:10]} {hh:02d}:{mm:02d}  종={r[5]}  누적량={r[6]}")
