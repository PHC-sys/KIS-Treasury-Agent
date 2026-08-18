# home/ — 로컬 작업본 스냅샷 (2026-08-18)

집 PC 로컬에서 작업한 결과를 **검토용으로** 통째로 올린 폴더다.
레포 루트(회사본)는 **건드리지 않았다**. 여기 내용을 확인한 뒤 루트에 반영하면 된다.

기준: 로컬 `KIS-Treasury-Agent-main/` 전체 복사.
제외한 것 — `db/`(원장·5분봉 sqlite), `raw/`, `data/`, `manual_input.csv`, `__pycache__`,
그리고 작업 중 만든 `*.BACKUP_*` 백업본과 엑셀 잠금파일(`~$*`).
**원장을 제외해도 무방하다**: `out/ktb_daily.csv` + `src/restore_ledger.py`로 완전 복원된다(아래 §4).

---

## 1. 무엇이 달라졌나

스키마 **42 → 47(MACRO) → 48(wti) → 56(수급 세분화) → 63(미국채 OHLC)**.
이 폴더는 **63컬럼** 기준이다.

### v2.6 — 선물 수급 세분화 8컬럼 (2026-08-15)
`sec_net_ktb3/10`(증권/선물) · `indiv_net_ktb3/10`(개인) · `etc_net_ktb3/10`(나머지 전부) · `pension_net_ktb3/10`(연기금)

- 인포맥스 FUT 투자자별 **순매수수량** 직접 수신(단위 계약, +매수/−매도).
- `etc_net` = 기타 + 국가/지방 + 연기금 + 종신금 (= "나머지 전부의 합").
- **`pension_net`은 `etc_net`의 부분집합**이다. 8개를 전부 더하면 연기금이 이중계산된다.
  → 실측상 `etc_net`의 86~88%가 연기금이라, '기타법인 흐름'으로 오독되지 않게 따로 노출한 것.
- **항등식 자동검증** 신설: `fx+bank+itrust+ins+sec+indiv+etc = 0` (`config.FLOW_IDENTITY`).
  KTB10 전 구간 불일치 0 / KTB3는 2004-06-28 이후 5,466일 불일치 0.
  그 이전은 원천이 세부분류를 안 줘서 검증 대상에서 제외(`FLOW_IDENTITY_START`).

### v2.7 — 미국채 2Y·10Y OHLC 7컬럼 (2026-08-18)
`ust10_open/high/low` + `ust2`, `ust2_open/high/low`. 기존 `ust10`은 **이름·값 그대로** 유지(= close 역할).

- 종목: `IR`/`US10Y`, `IR`/`US02Y`. 항목 `MID_Open/High/Low/Close`.
- **만기는 2자리 제로패딩** — `US02Y`·`US10Y`·`US20Y`·`US30Y`. `US2Y`·`US5Y` 등 1자리는 빈 응답.
- **`GVO:` 접두어 함정**: 단말 화면의 `GVO:TR10Y`/`GVO:TR02Y`는 시장구분이 **`FRN`**이다.
  `IR`로 넣으면 빈 응답이라 "미지원"으로 오판하기 쉽다. `FRN`의 `현재가` = `IR`의 `BID_Close`와 소수점까지 일치.
- **MID를 쓴다** — 기존 ust10 26년치가 MID_Close이기 때문. 단말 화면은 BID이라 0.0~0.5bp 높다.
  BID가 필요하면 같은 IR 호출의 `BID_*`를 **별도 컬럼**으로 추가할 것(기존 시리즈 변경 금지).

### 버그 수정 3건
| # | 버그 | 증상 |
|---|---|---|
| A | **as-of 동점 시 오래된 미국일 채택** | 한국 휴장 연휴 뒤 ust10·ust2·wti가 하루 낡은 값. 실측: 한국 2026-03-03 = 미국 02-27(3.942), 정답은 03-02(4.035) |
| B | **`trading_days` 전 구간 평일 폴백** | 한국 휴장일 286일을 거래일로 오인. `2026-05-26` ust10 유실 |
| C | **`_read_one`이 없는 항목에 `ValueError`** | 확장 전 워크북으로 덮어쓰면 FUT 시트 적재가 통째로 실패 → 그날 선물 데이터 전멸 |
| D | **ECOS 증분창 구멍** | cd91이 국고금리보다 먼저 공표되면 창이 닫혀 `y_ktb3/10/30`을 영영 안 긁음 (2026-08-12 실제 발생) |

원인·재발방지는 `docs/개발_인수인계.md` §7에 전부 기록했다. **특히 A는 조건이 까다로워 재발시키기 쉽다**:
`store.upsert`가 `executemany`라 나중 실행이 이기는데, 인포맥스 워크북이 `sort=D`(내림차순)이라
오래된 값이 나중에 덮어썼다. 폐기된 `fred.py`는 CSV가 오름차순이라 우연히 맞았던 것.
→ 이제 `_asof_records`가 한국일별 `max(미국일)`로 확정한다(입력 순서 무관).

---

## 2. 변경된 파일

| 파일 | 내용 |
|---|---|
| `src/config.py` | COLUMNS 63, SOURCE_FIELDS, `FLOW_IDENTITY`, `RECHECK_TRADING_DAYS`, `trading_days` 재작성 |
| `src/infomax_pull.py` | `_read_asof` 다중항목화, `_read_ust2`, `_asof_records` 동점처리, `_read_one` 방어, `("sum",…)` 매핑 |
| `src/build_infomax.py` | `_FUT` 23항목, `US10Y` 항목 확장, `US02Y` 시트 신설 |
| `src/update.py` | `flow_identity_report()`, `window_for` 재조회 룩백, 리포트 신규/정정 구분 |
| `src/restore_ledger.py` | **신규** — wide CSV → ledger 역복원 + 왕복검증 |
| `docs/개발_인수인계.md` | §7에 v2.6·v2.7·버그 A~D 전부 기록 |
| `CLAUDE.md` | 컬럼 수·항등식 주의사항 |
| `infomax_data.xlsx` | FUT3/FUT10 23항목, US10Y OHLC 확장, **US02Y 시트 신설**(각 9,000행 전이력) |
| `out/ktb_daily.csv` | **63컬럼 6,633행** (1999-09-29 ~ 2026-08-14) |

---

## 3. 반영 방법

### 방법 1 — 드롭인 (권장, 인포맥스 불필요)
```bash
# 레포 루트에서
cp -r home/src home/docs home/CLAUDE.md home/README.md home/requirements.txt .
cp home/infomax_data.xlsx home/out/ktb_daily.csv .   # 경로 맞춰서
python src\restore_ledger.py                          # CSV → db/ktb.sqlite 복원 + 자체검증
```
`infomax_data.xlsx`를 그대로 쓰면 인포맥스 재조회가 아예 필요 없다.

### 방법 2 — 코드만 가져가고 워크북은 직접 갱신
⚠**주의**: `US10Y` 항목 확장과 `FUT3`/`FUT10` 23항목은 **기존 시트 변경**이라
`build_infomax.py add`(없는 시트만 추가)로는 **반영되지 않는다**. 둘 중 하나를 해야 한다.
- `python src\build_infomax.py` (워크북 통째 재생성, 13시트 × 9000 — 10~20분)
- 또는 FUT3/FUT10/US10Y 시트만 표적 갱신하는 스크립트 작성

`US02Y`는 **새 시트**라 `ensure_sheets()`가 `run.bat` 실행 시 자동 추가한다(F1=9000 전이력).

COM 주의: 워크북을 **열면서** 셀을 쓰면 `RPC_E_CALL_REJECTED`가 난다.
`_refresh_wb`의 방어책(빈 워크북 확보 → 수동계산 → 유휴 대기 → `_com_retry`)을 따를 것.
실패 시 `wb.Close(SaveChanges=False)` — 반쯤 바뀐 시트를 저장하면 안 된다.

---

## 4. 검증

```bash
python src\restore_ledger.py --verify-only   # 원장 ↔ CSV 동등성 왕복검증
python src\update.py                          # 건강리포트 (수급항등식 포함)
```
정상이면:
```
[검증 1/2] ✅ 입력 CSV와 완전 일치 — ledger가 원본과 동등하다
[검증 2/2] ✅ 변화 없음
[수급항등식] 최근 20건 합=0  OK
```

---

## 5. 알려진 원천 특성 (버그 아님)

- **인포맥스 MID OHLC**에서 종가가 `[저가, 고가]` 밖인 날이 **35일/9,000일(0.4%)**, 최대 0.79bp, 2018년 이후 0건.
  `MID_Low=(BID_Low+ASK_Low)/2`라 bid·ask 최저점 시각이 다르면 발생. **보정하지 않고 원천 그대로 적재.**
- **일별 refresh의 이음매 소실**: `F1=120`으로 상위 120행만 덮어쓰는데 이음매 행 번호는 고정이라
  거래일이 하루 진행될 때마다 워크북에서 거래일 1개가 영구 소실된다(원장엔 남아 무해).
  분기 1회 `IMX_DAILY_COUNT=9000`으로 복구.
- **`real=false`가 "확정치만"이라는 뜻이 아니다**: `real=true`와 과거값 100% 동일하고
  진행 중인 당일 행만 다르다. 미확정 당일 값도 IMDH는 그대로 준다.
  우리 쪽은 `kdays`가 `ref_date()`까지만 만들어져 진행 중 미국일이 배정 대상 없이 스킵된다.
