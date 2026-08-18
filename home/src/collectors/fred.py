# -*- coding: utf-8 -*-
"""collectors/fred.py — FRED DGS10 → ust10 (한·미 시차 정합)

키 불필요 CSV 다운로드 엔드포인트(완전 자동, 등록 불필요). 결측은 '.'로 오므로 건너뛴다.

★한·미 시차(룩어헤드 방지): 미국 날짜 d의 종가는 그 다음 한국 거래일 새벽에야 확정된다.
  그래서 DGS10[미국날짜 d]를 'd 다음 한국 거래일'에 배정(as-of)한다. 그러면 한국 거래일 T의
  ust10 = T 직전 미국영업일 종가 = 그 세션에 실제로 알 수 있던 값(미래참조 제거).
  부수효과: 미국 공휴일(한국 개장)엔 직전 미국종가가 캐리돼 빈칸이 줄어든다.
"""
import bisect
import csv
import datetime as dt
import io

import requests

import config
from collectors.base import Collector, Record

URL = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10"


class FREDCollector(Collector):
    source = "fred"

    def fetch(self, start, end):
        r = requests.get(URL, timeout=30)
        r.raise_for_status()
        self.save_raw(r.text, "DGS10")

        # 배정 대상 한국 거래일 목록(경계 커버 위해 앞뒤로 여유).
        kdays = [k.isoformat() for k in
                 config.trading_days(start - dt.timedelta(days=15), end)]
        # 경계 커버: start보다 조금 앞선 미국날짜도 읽어 다음 한국거래일로 매핑.
        s = (start - dt.timedelta(days=10)).isoformat()
        e = end.isoformat()

        records = []
        rd = csv.reader(io.StringIO(r.text))
        next(rd, None)  # 헤더
        for row in rd:  # DGS10 CSV는 날짜 오름차순 → 같은 한국일에 겹치면 최신 미국일이 마지막에 이겨 upsert
            if len(row) < 2:
                continue
            d, v = row[0].strip(), row[1].strip()
            if v in ("", ".") or not (s <= d <= e):
                continue
            i = bisect.bisect_right(kdays, d)   # d 다음 한국 거래일(첫 T > d)
            if i >= len(kdays):
                continue                        # 아직 배정할 한국 거래일 없음(최신 미국일) → 다음 실행에서
            records.append(Record(date=kdays[i], field="ust10", value=float(v), as_of=d))
        return records
