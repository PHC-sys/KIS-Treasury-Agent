# -*- coding: utf-8 -*-
"""등록된 HTTP 수집기 목록. update.py가 이 순서로 돈다.
선물은 인포맥스 국채연결선물(infomax_pull.py)로 전부 처리 → KRX 미사용.
kofr도 인포맥스(ECO/785831). 인포맥스는 infomax_pull.py로 별도 적재."""
from collectors.ecos import ECOSCollector
from collectors.fred import FREDCollector
from collectors.fx import FXCollector

ACTIVE_COLLECTORS = [
    ECOSCollector(),
    FREDCollector(),
    FXCollector(),
]
