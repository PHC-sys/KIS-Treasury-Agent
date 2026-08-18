# KIS Treasury Agent — 데이터 수집 레이어

국채선물(3년·10년) 트레이딩 에이전트가 쓸 **일별 42컬럼 데이터**를 매일 수집·검증·적재하고
Supabase로 올려, 차장님(및 차장님 Claude)이 read-only로 어디서든 조회하게 하는 파이프라인.

- **이력**: 1999~현재, 거래일만 (`out/ktb_daily.csv`, Supabase `ktb_daily`)
- **확보**: 42컬럼 중 40 자동 (CDS=라이선스 미확보, note=사람 몫)

## 매일 하는 일 (워크플로우)

1. 아침에 **인포맥스 로그인**
2. **`run.bat`** 더블클릭 → 아래가 자동으로:
   ```
   [1/4] 인포맥스 엑셀(infomax_data.xlsx) 새로고침 (COM)
   [2/4] 인포맥스 → ledger 적재
   [3/4] 공개소스(ECOS·FRED) 수집 + 파생계산 + CSV export + 건강리포트
   [4/4] Supabase push
   ```

## 폴더 구조

```
KIS Treasury Agent/
├─ run.bat              🖱️ 진입점 (더블클릭 → src\run.py)
├─ README.md  CLAUDE.md  requirements.txt
├─ manual_input.csv     수기 입력 (kr_cds5 등 라이선스 자동불가분)
├─ infomax_data.xlsx    인포맥스 수집 워크북 (매일 refresh)
├─ src/                 ◀ 코드 전부 (서로 import → 함께 둠)
│   ├─ run.py            4단계 오케스트레이터
│   ├─ config.py         42컬럼 스키마·소스배정·sanity·기준일 — 단일 진실
│   ├─ update.py         공개소스 수집→sanity→upsert→파생→export→리포트
│   ├─ infomax_pull.py   인포맥스 엑셀 refresh(COM)/load (RegisterXLL로 IMDH)
│   ├─ build_infomax.py  워크북 생성(종목 C65/C67·항목 정의)
│   ├─ push_supabase.py  ktb_daily → Supabase upsert
│   ├─ setup_readonly.py 차장님용 read-only 롤 발급 (수동 실행)
│   ├─ sanity.py         적재 전 검증 (범위·동결·완결성)
│   ├─ store.py          ledger 접근 (멱등 upsert·wide export)
│   ├─ schema.sql        ledger/quarantine 스키마
│   └─ collectors/       ecos·fred·fx (HTTP 수집기)
├─ db/                  런타임 DB (ktb.sqlite = ledger, 진실)
├─ out/                 산출 CSV (ktb_daily.csv, 42컬럼 wide)
├─ raw/                 소스 원본 응답 보관
├─ data/               차장님 원본 체크리스트
├─ docs/               문서 (아래) + supabase_schema.sql(참고)
└─ Infomax Manual/     인포맥스 매뉴얼 PDF · 연결선물.xlsx · _archive_discovery/
```

## 데이터 흐름

```
소스 ──수집──▶ raw/(원본보관) ──sanity──▶ db/ktb.sqlite(ledger, 진실)
                                              │ + 파생계산(fx_cum·days_to_expiry·roll_flag·base_rate ffill)
                                              ▼
                                    out/ktb_daily.csv (42컬럼 wide)
                                              │
                                    Supabase Postgres (ktb_daily)
                                       ├ 우리: SUPABASE_DB_URL (write)
                                       └ 차장님/차장님 Claude: ktb_readonly (read-only)
```

**소스**: 선물 전부 = 인포맥스 국채연결선물(C65/C67) · 금리·환율 = ECOS · 미국채 = FRED · IRS/OIS = 인포맥스.
자세한 컬럼 설명은 **[docs/데이터_사전.md](docs/데이터_사전.md)**.

> **왜 코드가 전부 `src/`에 평평하게?** 파이썬 모듈들이 서로 `import config`·`import store`
> 형태로 참조한다. 실행은 `run.bat`이 `python src\run.py`로 돌리며, 이때 `src/`가
> 모듈 경로에 올라가 서로를 찾는다. (데이터 폴더 db/out/raw는 루트에 그대로.)

## 문서

| 문서 | 내용 |
|---|---|
| **[docs/데이터_사전.md](docs/데이터_사전.md)** | 각 컬럼이 뭔지·어디서 어떻게 뽑았나 (계획 대비 실제) |
| [docs/데이터수집_사양.md](docs/데이터수집_사양.md) | 원본 사양서 (Phase 0 설계) |
| [docs/STATUS.md](docs/STATUS.md) | 구축 진행 로그 (무슨 문제를 어떻게 풀었나) |
| [CLAUDE.md](CLAUDE.md) | 절대 규칙 (Claude Code용) |

## 환경변수 (사용자 PC)

| 변수 | 용도 |
|---|---|
| `ECOS_API_KEY` | 한국은행 ECOS (금리·환율) |
| `SUPABASE_DB_URL` | Supabase 쓰기 (Session pooler URI) |

## 차장님 read-only 접근

- **호스팅 MCP 링크**(설치 불필요): `https://mcp.supabase.com/mcp?project_ref=<ref>&read_only=true`
  → Claude 커스텀 커넥터에 붙이고 Supabase OAuth. 차장님 Claude가 DB 직접 조회.
- 또는 SQL 클라이언트에 `ktb_readonly` 접속문자열(`setup_readonly.py` 출력).
- 권한: **SELECT만** (write 원천 불가).

## 다음 단계

- 작업 스케줄러 야간 자동실행 (인포맥스 세션 유지 필요)
- 위층: 시그널 엔진 (이 42컬럼 → TREND/POSX/FLOW/VREG 점수화 → 매매신호)
