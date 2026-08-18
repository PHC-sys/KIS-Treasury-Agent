# -*- coding: utf-8 -*-
"""등록된 HTTP 수집기 목록. update.py가 이 순서로 돈다.
선물은 인포맥스 국채연결선물(infomax_pull.py)로 전부 처리 → KRX 미사용.
kofr도 인포맥스(ECO/785831). 인포맥스는 infomax_pull.py로 별도 적재.
★ust10: FRED(DGS10)는 공표 지연이 커서 인포맥스 IR/US10Y(MID_Close)로 이관 →
  FREDCollector 미사용(fred.py는 참고용 보존). infomax_pull._read_ust10이 처리."""
from collectors.ecos import ECOSCollector
from collectors.fx import FXCollector

ACTIVE_COLLECTORS = [
    ECOSCollector(),
    FXCollector(),
]
