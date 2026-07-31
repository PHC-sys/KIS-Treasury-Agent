# -*- coding: utf-8 -*-
"""
run.py — 아침 원버튼 일일 갱신 오케스트레이터

전제: 인포맥스에 로그인돼 있을 것(엑셀 애드인이 살아있어야 함).
순서:
  1) 인포맥스 엑셀 전부 새로고침(COM) — 종료일=오늘로 최신 종가까지
  2) 인포맥스 → ledger 적재
  3) 공개소스(ECOS·FRED·KRX) 수집 + manual 병합 + 파생계산 + CSV export + 건강리포트

인포맥스가 안 떠 있으면 1)만 건너뛰고 나머지는 정상 수행(공개소스는 항상 최신).
"""
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import infomax_pull
import update


def main():
    print("=" * 60)
    print("[1/4] 인포맥스 엑셀 새로고침 (COM)")
    try:
        infomax_pull.refresh_all()
    except Exception as e:
        print(f"  ⚠ 인포맥스 새로고침 실패(로그인/엑셀 확인) → 건너뜀: {e}")

    print("\n[2/4] 인포맥스 → ledger 적재")
    try:
        infomax_pull.load()
    except Exception as e:
        print(f"  ⚠ 인포맥스 적재 실패: {e}")

    print("\n[3/4] 공개소스 수집 + 파생 + export")
    update.main()          # ECOS·FRED + manual + 파생 + export + 건강리포트

    print("\n[4/4] Supabase push")
    try:
        import push_supabase
        push_supabase.main()
    except Exception as e:
        print(f"  ⚠ Supabase push 실패(SUPABASE_DB_URL/네트워크 확인) → 건너뜀: {e}")


if __name__ == "__main__":
    main()
