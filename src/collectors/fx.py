# -*- coding: utf-8 -*-
"""collectors/fx.py — USD/KRW (ECOS 731Y001 원/미국달러 매매기준율)

같은 ECOS 키 사용. 사양서상 FX는 별도 소스로 관리(스키마상 구분).
"""
import os
import requests

from collectors.base import Collector, Record

STAT, ITEM, CYC = "731Y001", "0000001", "D"
BASE = "https://ecos.bok.or.kr/api/StatisticSearch"


class FXCollector(Collector):
    source = "fx"

    def fetch(self, start, end):
        key = os.environ.get("ECOS_API_KEY")
        if not key:
            raise RuntimeError("ECOS_API_KEY 미설정")
        s, e = start.strftime("%Y%m%d"), end.strftime("%Y%m%d")
        url = f"{BASE}/{key}/json/kr/1/20000/{STAT}/{CYC}/{s}/{e}/{ITEM}"
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        self.save_raw(r.text, "usdkrw")
        j = r.json()
        records = []
        if "StatisticSearch" in j:
            for row in j["StatisticSearch"].get("row", []):
                v = (row.get("DATA_VALUE") or "").strip()
                t = (row.get("TIME") or "").strip()
                if not v or len(t) != 8:
                    continue
                d = f"{t[:4]}-{t[4:6]}-{t[6:]}"
                records.append(Record(date=d, field="usdkrw", value=float(v), as_of=d))
        return records
