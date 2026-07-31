-- KTB 일별 데이터 테이블 (Supabase SQL Editor에서 실행)
CREATE TABLE IF NOT EXISTS ktb_daily (
  date date PRIMARY KEY,
  ktb3_settle double precision,
  ktb3_chg double precision,
  ktb10_settle double precision,
  ktb10_chg double precision,
  ktb3_vol double precision,
  ktb10_vol double precision,
  ktb3_oi double precision,
  ktb10_oi double precision,
  fx_net_ktb3 double precision,
  fx_net_ktb10 double precision,
  fx_cum_ktb3 double precision,
  fx_cum_ktb10 double precision,
  bank_net_ktb3 double precision,
  itrust_net_ktb3 double precision,
  ins_net_ktb3 double precision,
  y_ktb3 double precision,
  y_ktb10 double precision,
  y_ktb30 double precision,
  base_rate double precision,
  cd91 double precision,
  kofr double precision,
  irs3y double precision,
  ois3y double precision,
  ust10 double precision,
  usdkrw double precision,
  kr_cds5 double precision,
  ktb3_theo_basis double precision,
  ktb10_theo_basis double precision,
  note text,
  ktb3_open double precision,
  ktb3_high double precision,
  ktb3_low double precision,
  ktb10_open double precision,
  ktb10_high double precision,
  ktb10_low double precision,
  days_to_expiry double precision,
  roll_flag double precision,
  bank_net_ktb10 double precision,
  itrust_net_ktb10 double precision,
  ins_net_ktb10 double precision,
  irs10y double precision
);

-- 읽기전용 롤 (차장님 & 차장님 클로드용)
CREATE ROLE ktb_readonly LOGIN PASSWORD 'CHANGE_ME_강한비번';
GRANT USAGE ON SCHEMA public TO ktb_readonly;
GRANT SELECT ON ALL TABLES IN SCHEMA public TO ktb_readonly;
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT SELECT TO ktb_readonly;