-- ledger: long-format 수집 원장 (사양서 §1)
-- 저장 키 = (date, field). 같은 키면 upsert(덮어쓰기) → 멱등성.
-- 값이 결측이면 행 자체를 안 만든다 (빈칸 = 행 없음).
CREATE TABLE IF NOT EXISTS ledger (
  date       TEXT NOT NULL,   -- 'YYYY-MM-DD' 데이터 기준 영업일
  field      TEXT NOT NULL,   -- ktb_daily 컬럼명
  value      REAL,            -- 실제 값
  source     TEXT NOT NULL,   -- 'krx'|'ecos'|'fred'|'kofr'|'fx'|'manual'|'derived'
  as_of      TEXT NOT NULL,   -- 소스가 밝힌 데이터 기준일
  fetched_at TEXT NOT NULL,   -- 인출 시각 (ISO8601)
  PRIMARY KEY (date, field)
);

CREATE INDEX IF NOT EXISTS idx_ledger_source ON ledger(source);
CREATE INDEX IF NOT EXISTS idx_ledger_field  ON ledger(field);

-- sanity 실패 격리 로그 (조용히 넘어가지 않는다 — 사양서 §5)
CREATE TABLE IF NOT EXISTS quarantine (
  date       TEXT NOT NULL,
  field      TEXT NOT NULL,
  value      REAL,
  source     TEXT NOT NULL,
  reason     TEXT NOT NULL,   -- 어떤 sanity에서 걸렸나
  fetched_at TEXT NOT NULL,
  PRIMARY KEY (date, field, fetched_at)
);
