# -*- coding: utf-8 -*-
"""인트라데이 날짜범위 제어 크랙: B1/D1에 시각까지 넣어 07-31만 오는지."""
import win32com.client as win32
import pythoncom, time, sys, datetime as dt
sys.path.insert(0, r"D:\KIS Treasury Agent\src")
import infomax_pull
def test(b1, d1, label):
    pythoncom.CoInitialize()
    app=win32.Dispatch("Excel.Application")
    for _ in range(15):
        try: app.Version; break
        except: time.sleep(1)
    app.DisplayAlerts=False; app.Visible=False
    infomax_pull._register_addin(app)
    wb=app.Workbooks.Add(); ws=wb.Worksheets(1)
    ws.Range("B1").Value=b1; ws.Range("D1").Value=d1; ws.Range("F1").Value=99999
    for j,it in enumerate(["일자","시간","시가","고가","저가","현재가","거래량","미결제약정수량"]): ws.Cells(3,j+1).Value=it
    opt="Per=MM,Cycle=5,sort=D,real=false,Bizday=0,Quote=종가,Pos=20,Orient=V,Title=T,DtFmt=1,TmFmt=1,unit=true"
    ws.Range("A2").Formula=f'=IMDH("FUT","C65",A3:H3,$B$1,$D$1,$F$1,"{opt}")'
    time.sleep(13)
    data=ws.Range("A4:H700").Value
    rows=[r for r in data if r[0] is not None]
    dates=sorted(set(str(r[0])[:10] for r in rows))
    print(f"[{label}] 봉수={len(rows)}, 날짜={dates[:2]}..{dates[-2:]}")
    d31=[r for r in rows if str(r[0])[:10]=="2026-07-31"]
    if d31:
        hs=[float(r[3]) for r in d31 if r[3] not in(None,'')]; ls=[float(r[4]) for r in d31 if r[4] not in(None,'')]
        print(f"   07-31 {len(d31)}봉: 최신봉현재가={d31[0][5]}(vs 103.23) 고={max(hs)}(103.25) 저={min(ls)}(103.06) OI={d31[0][7]}(597425)")
    try: wb.Close(SaveChanges=False)
    except: pass
    try: app.Quit()
    except: pass
test(dt.datetime(2026,7,31,9,0), dt.datetime(2026,7,31,15,45), "07-31 09:00~15:45")
