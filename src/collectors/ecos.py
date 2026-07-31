# -*- coding: utf-8 -*-
"""collectors/ecos.py — 한국은행 ECOS API

담당: base_rate, cd91, kofr, y_ktb3, y_ktb10, y_ktb30
인증키는 환경변수 ECOS_API_KEY (하드코딩 금지). 실측 완료된 통계·항목 코드 사용.
"""
import os
import requests

from collectors.base import Collector, Record

# field → (통계표코드, 항목코드, 주기)  ※ 전부 일별(D). base_rate만 캘린더 매일값.
FIELD_MAP = {
    "base_rate": ("722Y001", "0101000", "D"),
    "cd91":      ("817Y002", "010502000", "D"),
    "kofr":      ("817Y002", "010901000", "D"),
    "y_ktb3":    ("817Y002", "010200000", "D"),
    "y_ktb10":   ("817Y002", "010210000", "D"),
    "y_ktb30":   ("817Y002", "010230000", "D"),
}
BASE = "https://ecos.bok.or.kr/api/StatisticSearch"


class ECOSCollector(Collector):
    source = "ecos"

    def fetch(self, start, end):
        key = os.environ.get("ECOS_API_KEY")
        if not key:
            # 실패는 시끄럽게 — 키 없으면 빈 값을 성공인 척 반환하지 않는다.
            raise RuntimeError("ECOS_API_KEY 미설정")

        s, e = start.strftime("%Y%m%d"), end.strftime("%Y%m%d")
        records = []
        for field, (stat, item, cyc) in FIELD_MAP.items():
            url = f"{BASE}/{key}/json/kr/1/20000/{stat}/{cyc}/{s}/{e}/{item}"
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            self.save_raw(r.text, field)
            j = r.json()
            if "StatisticSearch" not in j:
                # RESULT 코드가 오면(INFO-200=해당없음 등) 그 필드만 건너뜀
                continue
            for row in j["StatisticSearch"].get("row", []):
                v = (row.get("DATA_VALUE") or "").strip()
                t = (row.get("TIME") or "").strip()
                if not v or len(t) != 8:
                    continue
                d = f"{t[:4]}-{t[4:6]}-{t[6:]}"
                records.append(Record(date=d, field=field, value=float(v), as_of=d))
        return records
