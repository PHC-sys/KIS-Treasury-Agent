# -*- coding: utf-8 -*-
"""collectors/fred.py — FRED DGS10 → ust10

키 불필요 CSV 다운로드 엔드포인트(완전 자동, 등록 불필요). 실측 완료.
결측은 '.'로 오므로 건너뛴다.
"""
import csv
import io
import requests

from collectors.base import Collector, Record

URL = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10"


class FREDCollector(Collector):
    source = "fred"

    def fetch(self, start, end):
        r = requests.get(URL, timeout=30)
        r.raise_for_status()
        self.save_raw(r.text, "DGS10")
        s, e = start.isoformat(), end.isoformat()
        records = []
        rd = csv.reader(io.StringIO(r.text))
        next(rd, None)  # 헤더
        for row in rd:
            if len(row) < 2:
                continue
            d, v = row[0].strip(), row[1].strip()
            if v in ("", ".") or not (s <= d <= e):
                continue
            records.append(Record(date=d, field="ust10", value=float(v), as_of=d))
        return records
