# -*- coding: utf-8 -*-
"""
restore_ledger.py — wide CSV → db/ktb.sqlite(ledger) 역복원

ledger(db/ktb.sqlite)는 .gitignore 제외라 새 PC엔 없다. 반면 wide 산출물
(out/ktb_daily.csv 또는 Supabase ktb_daily 내보내기)은 ledger의 완전한 투영이므로
거꾸로 melt 하면 ledger를 되살릴 수 있다. 그게 이 스크립트다.

    python src\\restore_ledger.py                       # out/ktb_daily.csv → db/ktb.sqlite
    python src\\restore_ledger.py --csv supabase.csv    # Supabase 내보내기에서
    python src\\restore_ledger.py --verify-only         # DB 안 건드리고 왕복검증만
    python src\\restore_ledger.py --force               # 기존 ktb.sqlite 덮어쓰기(백업 후)

복원되는 것 / 안 되는 것
  ✓ (date, field, value)            — 원장의 실체. 100% 복원.
  ✓ source                          — config.SOURCE_FIELDS가 필드→소스 유일매핑이라 역산 가능.
                                      (증분창 판단 last_date_for_source가 이걸 쓴다)
  △ as_of                           — wide CSV에 없음 → date로 채움. 파이프라인 로직은 as_of를
                                      읽지 않으므로(메타 전용) 무해. 단 ust10의 as_of는 원래
                                      '미국일자'였는데 여기선 한국일자가 된다. 값·배정은 그대로.
  △ fetched_at                      — 복원 시각으로 기록.
  ✗ 휴장일 행                        — wide CSV는 거래일만 담는다. base_rate 캘린더-매일 값 등
                                      비거래일 행은 유실되나, export가 어차피 거래일만 쓰고
                                      base_rate는 recompute_derived가 ffill로 되살린다.
  ✗ quarantine 이력                  — 유실(재현 불가). 실사용 영향 없음.

끝나면 왕복검증을 한다: 복원 → recompute_derived → 임시경로로 export → 입력 CSV와 바이트 비교.
일치하면 ledger가 원본과 동등하다는 뜻이다. (검증 중 실제 out/ktb_daily.csv는 건드리지 않는다)
"""
import argparse
import csv
import shutil
import sys
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import config


# ── 필드 → 소스 역산 (config가 단일 진실) ────────────────────────────
def field_source_map():
    """{field: source}. SOURCE_FIELDS는 필드당 정확히 한 소스를 보장(config의 assert)."""
    m = {f: src for src, fields in config.SOURCE_FIELDS.items() for f in fields}
    for f in config.DERIVED_FIELDS:
        m[f] = "derived"
    return m


def ecos_frontier(rows):
    """ECOS '실제 공표' 최종일 = base_rate 제외한 ecos 필드들의 마지막 날짜.

    base_rate는 recompute_derived가 금통위까지 ffill 하므로 최신 날짜까지 값이 있다.
    이걸 그대로 source='ecos'로 넣으면 last_date_for_source('ecos')가 미래로 밀려
    아직 미공표였던 cd91·국고금리를 영영 안 긁는다(store.py 주석 참조).
    → frontier 이후의 base_rate는 source='derived'로 표시해 증분창을 오염시키지 않는다."""
    real = [f for f in config.SOURCE_FIELDS["ecos"] if f != "base_rate"]
    dates = [r["date"] for r in rows if any(r.get(f) for f in real)]
    return max(dates) if dates else None


# ── 복원 ─────────────────────────────────────────────────────────────
def read_wide(csv_path):
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        raise SystemExit(f"입력 CSV가 비었다: {csv_path}")
    have = set(rows[0].keys())
    if "date" not in have:
        raise SystemExit(f"'date' 컬럼이 없다 — wide CSV가 맞나? {csv_path}")
    missing = [c for c in config.COLUMNS if c not in have]
    extra = [c for c in have if c not in config.COLUMNS]
    if missing:
        print(f"  ⚠ 입력에 없는 스키마 컬럼 {len(missing)}개(빈칸 처리): {', '.join(missing)}")
    if extra:
        print(f"  ⚠ 스키마에 없는 입력 컬럼(무시): {', '.join(extra)}")
    return rows


def to_ledger_rows(rows):
    """wide dict 리스트 → [(date, field, value, source, as_of, fetched_at), ...]"""
    src_of = field_source_map()
    frontier = ecos_frontier(rows)
    fetched = datetime.now().astimezone().isoformat(timespec="seconds")
    out, per_source, skipped = [], {}, 0
    for r in rows:
        d = (r.get("date") or "").strip()
        if len(d) != 10:
            continue
        for f in config.DATA_FIELDS:          # date/note 제외
            raw = (r.get(f) or "").strip()
            if raw == "":
                continue                       # 빈칸 = 행 없음 (규율 §6)
            try:
                v = float(raw)
            except ValueError:
                skipped += 1
                continue
            src = src_of[f]
            if f == "base_rate" and frontier and d > frontier:
                src = "derived"                # ffill 꼬리 — 증분창 오염 방지
            out.append((d, f, v, src, d, fetched))
            per_source[src] = per_source.get(src, 0) + 1
    if skipped:
        print(f"  ⚠ 숫자로 못 읽어 건너뛴 셀 {skipped}개")
    return out, per_source, frontier


def restore(csv_path, force):
    import store

    db = config.DB_PATH
    if db.exists():
        if not force:
            raise SystemExit(
                f"이미 존재: {db}\n"
                f"  덮어쓰려면 --force (기존 파일은 .bak_<타임스탬프>로 백업된다)")
        bak = db.with_suffix(f".sqlite.bak_{datetime.now():%Y%m%d_%H%M%S}")
        shutil.move(str(db), str(bak))
        print(f"  기존 DB 백업 → {bak.name}")

    rows = read_wide(csv_path)
    print(f"입력: {csv_path}  ({len(rows)}행, {rows[0]['date']} ~ {rows[-1]['date']})")
    ledger, per_source, frontier = to_ledger_rows(rows)
    print(f"  ECOS 공표 최종일(frontier) = {frontier}  → 이후 base_rate는 derived로 표기")

    conn = store.connect()
    store.init_db(conn)
    store.upsert(conn, ledger)
    print(f"\n[복원] ledger {len(ledger):,}행 → {db}")
    for s in sorted(per_source):
        print(f"    {s:8} {per_source[s]:>8,}")
    return conn


# ── 왕복검증: 복원본에서 다시 export → 입력 CSV와 비교 ────────────────
TMP_EXPORT = "_verify_export.csv"


def _export_temp(conn):
    """임시경로로 export. 실제 out/ktb_daily.csv는 절대 건드리지 않는다."""
    import update
    tmp = config.OUT_DIR / TMP_EXPORT
    real = config.OUT_CSV
    config.OUT_CSV = tmp
    try:
        n = update.export_csv(conn)
    finally:
        config.OUT_CSV = real
    return tmp, n


def _diff_cells(path_a, path_b):
    """두 wide CSV를 (날짜, 컬럼) 셀 단위로 비교. 반환 (헤더불일치, 날짜차, 셀차 리스트)."""
    ra = list(csv.reader(open(path_a, newline="", encoding="utf-8-sig")))
    rb = list(csv.reader(open(path_b, newline="", encoding="utf-8")))
    if ra[0] != rb[0]:
        return True, [], []
    hdr = ra[0]
    da = {r[0]: r for r in ra[1:]}
    db_ = {r[0]: r for r in rb[1:]}
    only = sorted(set(da) ^ set(db_))
    cells = []
    for d in sorted(set(da) & set(db_)):
        for i, (x, y) in enumerate(zip(da[d], db_[d])):
            if x != y:
                cells.append((d, hdr[i], x, y))
    return False, only, cells


def verify(conn, csv_path):
    """2단계 검증.
      1) 순수 복원 재현성 — 복원 직후 export가 입력 CSV와 바이트 동일해야 한다.
      2) 파생 재계산 델타  — recompute_derived 후 무엇이 바뀌는지 미리 보여준다.
         (wide CSV엔 거래일 행만 있어 원장의 '비거래일 행'이 유실된다. roll_flag는
          days_to_expiry의 전일 대비 증가로 정의돼 그 유실에 영향을 받는다 —
          롤 직후가 휴장일이었던 소수 날짜에서 0→1로 바뀐다. run.bat도 똑같이 바꾼다.)
    반환: 1)이 통과했는지."""
    import update

    tmp, n = _export_temp(conn)
    bad_hdr, only, cells = _diff_cells(csv_path, tmp)
    print(f"\n[검증 1/2] 순수 복원 재현성 — 재export {n}행")
    ok = not (bad_hdr or only or cells)
    if bad_hdr:
        print("  ✗ 헤더 불일치 — 입력 CSV 스키마가 config.COLUMNS와 다르다")
    elif only:
        print(f"  ✗ 한쪽에만 있는 날짜 {len(only)}개: {', '.join(only[:10])}")
    elif cells:
        print(f"  ✗ 값이 다른 셀 {len(cells)}개")
        for d, c, x, y in cells[:10]:
            print(f"      {d} {c}: 입력={x!r} 재export={y!r}")
    else:
        print("  ✅ 입력 CSV와 완전 일치 — ledger가 원본과 동등하다")

    update.recompute_derived(conn)
    tmp, n = _export_temp(conn)
    _, only2, cells2 = _diff_cells(csv_path, tmp)
    print(f"\n[검증 2/2] 파생 재계산 후 델타 — run.bat이 바꿀 값 미리보기")
    if not (only2 or cells2):
        print("  ✅ 변화 없음")
    else:
        by_col = {}
        for d, c, x, y in cells2:
            by_col.setdefault(c, []).append((d, x, y))
        for c, items in sorted(by_col.items(), key=lambda kv: -len(kv[1])):
            head = ", ".join(f"{d}({x}→{y})" for d, x, y in items[:6])
            more = f" 외 {len(items)-6}건" if len(items) > 6 else ""
            print(f"  ~ {c}: {len(items)}건 — {head}{more}")
        if only2:
            print(f"  ~ 날짜 증감 {len(only2)}개: {', '.join(only2[:10])}")
        print("  (파생 필드만 바뀌면 정상 — 비거래일 행 유실의 알려진 부작용)")

    (config.OUT_DIR / TMP_EXPORT).unlink(missing_ok=True)
    return ok


def main():
    ap = argparse.ArgumentParser(description="wide CSV → ledger 역복원")
    ap.add_argument("--csv", default=str(config.OUT_CSV),
                    help="입력 wide CSV (기본: out/ktb_daily.csv)")
    ap.add_argument("--db", help="출력 sqlite 경로 (기본: db/ktb.sqlite)")
    ap.add_argument("--force", action="store_true", help="기존 DB 백업 후 덮어쓰기")
    ap.add_argument("--verify-only", action="store_true",
                    help="복원 없이, 기존 DB로 왕복검증만")
    args = ap.parse_args()

    if args.db:
        config.DB_PATH = Path(args.db)

    if args.verify_only:
        import store
        if not config.DB_PATH.exists():
            raise SystemExit(f"DB 없음: {config.DB_PATH}")
        conn = store.connect()
        ok = verify(conn, args.csv)
        conn.close()
        return 0 if ok else 1

    conn = restore(args.csv, args.force)
    ok = verify(conn, args.csv)
    conn.close()
    if ok:
        print("\n다음 단계: 인포맥스 로그인 → run.bat (또는 python src\\update.py)")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
