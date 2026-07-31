# -*- coding: utf-8 -*-
"""setup_readonly.py — 차장님용 읽기전용(SELECT) 롤 생성 + 연결문자열 출력.

admin(SUPABASE_DB_URL)로 접속해 ktb_readonly 롤 생성/갱신.
출력되는 read-only 연결문자열을 차장님께 전달(그 클로드가 Postgres MCP로 붙음).

주의: 실행할 때마다 read-only 비밀번호를 새로 발급(회전)한다. 이미 차장님께
공유한 연결문자열이 있으면 재실행 후 새 문자열로 교체해야 한다.
수동 셋업 스크립트라 daily run.bat에는 포함되지 않는다.  ★import만으로는
아무 일도 안 일어나게 반드시 `python src/setup_readonly.py`로 직접 실행할 것."""
import os
import secrets
import time
from urllib.parse import urlparse, quote

import psycopg2


def main():
    url = os.environ["SUPABASE_DB_URL"]
    p = urlparse(url)
    role, ref = p.username.split(".", 1)          # postgres.<projectref>
    host, port = p.hostname, p.port or 5432
    ro_pw = secrets.token_urlsafe(18)             # URL-safe 강한 비번

    conn = psycopg2.connect(url)
    conn.autocommit = True
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM pg_roles WHERE rolname='ktb_readonly'")
    if cur.fetchone():
        cur.execute("ALTER ROLE ktb_readonly LOGIN PASSWORD %s", (ro_pw,))
    else:
        cur.execute("CREATE ROLE ktb_readonly LOGIN PASSWORD %s", (ro_pw,))
    cur.execute("GRANT USAGE ON SCHEMA public TO ktb_readonly")
    cur.execute("GRANT SELECT ON ALL TABLES IN SCHEMA public TO ktb_readonly")
    cur.execute("ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT SELECT ON TABLES TO ktb_readonly")
    conn.close()

    ro_uri = (f"postgresql://ktb_readonly.{ref}:{quote(ro_pw)}@{host}:{port}/postgres")
    print("읽기전용 롤 생성 완료.\n")
    print("=== 차장님 전달용 read-only 연결문자열 ===")
    print(ro_uri)
    print("\n(SELECT만 허용. INSERT/UPDATE/DELETE 원천 불가)")

    # 검증: 읽기전용으로 붙어서 SELECT 되고 INSERT 막히는지.
    # 세션 풀러(Supavisor)가 바뀐 비번을 전파하는 데 몇 초 걸릴 수 있어 재시도.
    print("\n=== 검증 ===")
    rc = None
    for attempt in range(1, 6):
        try:
            rc = psycopg2.connect(ro_uri)
            break
        except psycopg2.OperationalError as e:
            if attempt == 5:
                raise
            print(f"  풀러 전파 대기… ({attempt}/5) {str(e).splitlines()[0][:40]}")
            time.sleep(3)
    rcur = rc.cursor()
    rcur.execute("SELECT count(*) FROM ktb_daily")
    print("SELECT ok:", rcur.fetchone()[0], "행")
    try:
        rcur.execute("INSERT INTO ktb_daily(date) VALUES ('1900-01-01')")
        rc.commit()
        print("INSERT ⚠ 됨 (권한 확인 필요!)")
    except Exception as e:
        rc.rollback()
        print("INSERT 차단 확인 OK:", str(e).splitlines()[0][:50])
    rc.close()


if __name__ == "__main__":
    main()
