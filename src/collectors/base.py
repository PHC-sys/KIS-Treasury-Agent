# -*- coding: utf-8 -*-
"""
collectors/base.py — 수집기 공통 인터페이스

각 수집기는 fetch(start, end) → list[Record] 만 책임진다.
sanity·upsert·raw저장·파생계산은 전부 update.py가 한다 (수집기는 순수).
"""
from dataclasses import dataclass


@dataclass(frozen=True)
class Record:
    date: str       # 'YYYY-MM-DD' 데이터 기준 영업일
    field: str      # ktb_daily 컬럼명
    value: float
    as_of: str      # 소스가 밝힌 기준일 (모르면 date와 동일하게)


class Collector:
    """모든 수집기의 부모. source 이름과 fetch만 구현하면 된다."""
    source = "base"

    def fetch(self, start, end):
        """[start, end] 구간(포함)의 Record 리스트 반환.
        - 신규 없음 → 빈 리스트 (에러 아님).
        - 결측 필드 → Record를 만들지 않는다 (빈칸 = 행 없음).
        - 값이 의심스러워도 수집기는 판단하지 않는다. sanity가 판단.
        - 실패는 시끄럽게: 예외를 삼키지 말고 올린다.
        """
        raise NotImplementedError

    def save_raw(self, raw_bytes, tag):
        """원본 응답을 raw/ 에 타임스탬프로 보관 (재구성용, 수정 금지)."""
        from datetime import datetime
        import config
        config.RAW_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = config.RAW_DIR / f"{self.source}_{tag}_{stamp}.raw"
        mode = "wb" if isinstance(raw_bytes, (bytes, bytearray)) else "w"
        with open(path, mode, encoding=None if "b" in mode else "utf-8") as f:
            f.write(raw_bytes)
        return path
