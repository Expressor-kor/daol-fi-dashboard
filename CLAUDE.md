# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> 이 파일은 Claude Code가 이 저장소에서 **매 세션 자동으로 읽는 상시 운영 지침**이다.
> 일회성 작업 지시가 아니라, 이 코드베이스에서 일할 때 **항상 지켜야 할 계약·지뢰·거버넌스**를 담는다.
> 상세 이력·결정 근거는 `DAOL_FI_DASHBOARD_CHARTER_v2_1.md`, 진행 중 작업은 해당 작업 인수인계서 참조.

---

## 1. 이 프로젝트가 무엇인가 (북극성)

DAOL 채권운용본부 **주간전략회의 의사결정을 DB화**하는 라이브 HTML 대시보드. 11명의 FI 멤버가 매주 금리/커브/크레딧 전략 의견을 입력하고, 성과를 사후 평가한다.

```
[1단계] 주간 의사결정 DB화        ← 지금 (운영·완성 단계)
[2단계] 시장 국면 분석 + 성과 분석 모듈
[3단계] AI 종합 분석층
[4단계] AI 기반 모델 포트폴리오 제안 (사람이 최종 결정)
```

**1단계 성공 기준 = "훗날 AI가 읽어도 신뢰할 수 있는 깨끗한 의사결정 원장(ledger)".**
→ **화면 편의성과 데이터 무결성이 충돌하면 무결성이 이긴다.** 이 한 문장이 모든 판단의 기준이다.

작업 성격: 신규 구축이 아니라 **버그 수정 + 완성도 향상.**

---

## 2. 사용자 응대 방식

- 국내 채권 펀드매니저(2007~), 크레딧 애널리스트 경험, 경영학과 출신, **코딩 비전공자.**
- 금융 용어(KRD/IR/TE/듀레이션/커브/스프레드/BM 상대 듀레이션/attribution)는 자연스럽게 통한다 — 같은 수준으로 응대.
- 숫자·가격 **정확성**을 중시한다. 추정으로 숫자를 지어내지 말고, 모르면 모른다고 한다.
- 코드는 직접 안 짠다. **무엇을 왜 바꾸는지 평이하게 설명**하고, 트레이드오프(단점 포함)를 솔직하게 제시한다. 구조적 문서화를 선호한다.

---

## 3. 플랫폼 · 작업 방식

- **체제**: Azure Static Web Apps + GitHub 자동 배포. Microsoft Graph API로 SharePoint 직접 read/write. (Power Apps는 영구 폐기 — 재검토 없음.)
- **코드 작업**: **VS Code + Claude Code.** Claude Code가 로컬 `app.js`를 직접 읽고 수정 → 사용자 검토 → Source Control에서 커밋·푸시.
- **배포 절차**: 수정 → 사용자 검토 → 커밋·푸시 → GitHub Actions(`main` 브랜치 push 시 자동 트리거) → Azure Static Web Apps 배포 완료(초록불 ✓) → 브라우저 **Ctrl+F5**.
- **빌드 도구 없음**: `package.json`·`node_modules`·빌드 스텝이 없다. 순수 바닐라 JS 정적 파일. `npm install` / lint / 테스트 러너 없음.
- **문법 검사**: `node --check app.js` (Node.js가 PATH에 있을 때). 그 외에는 브라우저 DevTools 콘솔(F12)에서 직접 확인.
- **로컬 개발**: MSAL 인증은 `redirectUri: "auto"`로 동작하며, 로컬 파일(`file://`) 열기는 Graph API 호출 불가. 실질적 테스트는 배포 후 라이브 URL에서 진행.
- **인증**: MSAL `loginPopup`, 권한 `User.Read` + `Sites.ReadWrite.All`.

---

## 4. 데이터 계약 — 불변 원칙 (협상 불가 · 어기는 저장 금지)

이 장은 플랫폼·코드가 바뀌어도 변하지 않는다. 위반 소지가 보이면 **멈추고 사용자에게 보고**한다.

1. **`week + memberId`는 유일하다.** 저장은 upsert(있으면 PATCH, 없으면 POST).
2. **생략은 0이 아니라 빈 값(null)이다.** `null`/`Blank` 저장. 0 강제 금지.
3. **`0`과 `생략`은 서로 다른 값이다.**
   - `0` = "의견 있음, 중립/플랫 베팅" → 평균·중앙값 **포함**.
   - `생략`(null) = "의견 없음/미입력" → 집계에서 **제외**.
   - 집계 함수(`avg`/`median`)는 이미 `null`을 제외함. 평균이 이상하면 원인은 함수가 아니라 **생략이 0으로 저장된 데이터**다.
4. **판단 근거는 전략별 분리 저장**: `rateRationaleText` / `curveRationaleText` / `creditRationaleText`.
5. **확신도는 전략별 저장**: `rateConfidence` / `curveConfidence` / `creditConfidence`.
6. **부호 체계를 절대 반대로 받지 않는다.** (5장) 입력 화면에 항상 노출.
7. **`memberId`는 변경하지 않는다.** 삭제는 `active=false`.
8. **운영 테이블과 마감본(Archive) 분리.** 시장 기준일 변경은 이력(`FI_MarketMappingHistory`)을 남긴다.

---

## 5. 전략 부호 체계 (입력 화면 필수 노출 · 변경 불가)

| 전략 | key | `+` 의미 | `-` 의미 |
|---|---|---|---|
| 금리 방향성 | `rate` | 금리 하락 베팅 | 금리 상승 베팅 |
| 국고 3Y-10Y 커브 | `curve` | 스프레드 축소 | 스프레드 확대 |
| 여전 AA- 2Y − 통안 2Y | `credit` | 스프레드 축소 | 스프레드 확대 |

---

## 6. 성과 계산식 (드리프트 금지)

평가 매칭: **Wn 의견 → Wn~Wn+1 시장 변화로 평가.**

```
금리   손익 = -2.85 × rateDuration   × 국고3년 금리 변화폭(bp)
커브   손익 = -8.00 × curveDuration  × 3Y10Y 스프레드 변화폭(bp)
크레딧 손익 = -1.85 × creditDuration × creditSpread 변화폭(bp)
```

**검산 (구현·수정 후 반드시 일치 확인):**
```
금리   +0.10 / 국고3년 -10bp       → +2.85bp
커브   -0.10 / 3Y10Y  +5bp         → +4.00bp
크레딧 +0.10 / creditSpread -5bp   → +0.925bp
```

---

## 7. 컬럼명 지뢰 — 저장 코드가 반드시 준수

### (A) 이름이 다른 칸 — 올바른 내부명으로 보낼 것
| 리스트 | 잘못된 이름 ❌ | 실제 내부 컬럼명 ✅ |
|---|---|---|
| `FI_MarketData` | `creditAA2y` | **`creditAAm2y`** |
| `FI_WeeklyOpinions` | `rateRangeLow/High` | **`ktb3yRangeLow` / `ktb3yRangeHigh`** |
| `FI_WeeklyOpinions` | `curveRangeLow/High` | **`curveSpreadRangeLow` / `curveSpreadRangeHigh`** |
| `FI_WeeklyOpinions` | `creditSpreadRangeLow/High` | **`creditRangeLow` / `creditRangeHigh`** |
| `FI_WeeklySummaries` | `week` | **`Title`** (= week 키) |

### (B) 리스트에 없는 레거시 칸 — 절대 전송 금지 (보내면 graph-400 → 저장 전체 실패)
| 레거시 칸 ❌ | 대체 (실제 사용 칸) ✅ |
|---|---|
| `confidence` (단일) | `rateConfidence` / `curveConfidence` / `creditConfidence` |
| `rationaleText` (단일) | `rateRationaleText` / `curveRationaleText` / `creditRationaleText` |
| `rationaleTags` | (미사용 — 보내지 않음) |

> 저장 안전장치: `omitSharePointFields`가 실제 리스트에 없는 칸을 전송 전 제거한다. 저장 함수마다 이 점검이 들어가야 한다.
> 현재 omit 목록: `["confidence", "rationaleTags", "rationaleText"]` — 이 3개 외 칸을 추가·제거할 때는 사용자 명시 승인 필수.

---

## 8. 아키텍처 실측 (헷갈리기 쉬운 지점)

- **금리 그래프 원천 = SharePoint 문서 라이브러리 `rates_raw.xlsx`** (`state.dailyRates`). `fetchSharePointRatesBlob`이 Graph API로 `Shared Documents/주간전략회의/rates_raw.xlsx`를 다운로드 → `parseRatesWorkbook`(자체 xlsx 파서, 헤더 앵커 방식)이 파싱 → `normalizeDailyRate`가 `curveSpread`(`10Y−3Y`)·`creditSpread`(`AA-2Y−통안2Y`)를 자동 산출. `FI_DailyRates` 리스트는 `loadSharePointState`에서 아직 읽히나 직후 문서 라이브러리 파일이 **덮어쓰므로 사실상 은퇴 대기**(C4에서 CSV 백업 후 제거 예정).
- **주차-기준일 매핑은 원천과 독립**: `FI_MarketWeekMappings.baseDate` → `date === baseDate`인 금리 행이 그 주차 스냅샷. 원천을 바꿔도 매핑 메커니즘은 그대로.
- **컬럼명 자동 해석기** `resolveSharePointFieldNameFromSources`가 표시명/샘플키로 실제 내부 컬럼명을 찾아 매핑한다.
- **상태 이중 저장**: `state` 전역 객체가 유일한 런타임 진실 원천. `saveState()`는 `localStorage`(키 `daol-fi-strategy-dashboard-v2`)에만 쓴다. `loadSharePointState()`는 SharePoint에서 읽어 `state`를 덮어쓴 뒤 `saveState()` 호출. 두 원천이 다를 땐 SharePoint가 이긴다.
- **`render()`는 유일한 UI 갱신 진입점**: 상태 변경 후 반드시 `render()`를 호출해야 화면이 바뀐다. 개별 `render*()` 함수를 직접 호출하는 건 일부 갱신만 일어남에 주의.
- **`els` 객체**: 모든 DOM 요소가 파일 상단(약 220~285행)의 `const els = {...}`에 일괄 캐시된다. HTML에 요소를 추가하면 `els`에도 반드시 추가해야 한다.
- **`OPERATING_WEEK_CUTOFF = "2026-W24"`**: 이 상수보다 이른 주차는 운영 데이터로 취급하지 않는 필터(`filterOperatingWeeks`)가 걸려 있다. 데이터가 안 보일 때 이 상수를 먼저 확인.

### `app.js` 구역 지도 (약 4,980행)
| 행 범위 | 내용 |
|---|---|
| 1 – 220 | 전역 상수·설정 (`MICROSOFT_CONFIG`, `STRATEGIES`, Archive 스키마 등) |
| 220 – 290 | DOM 캐시(`els`), 전역 변수(`state`, `selectedWeek` 등) |
| 290 – 430 | 로컬 상태 관리 (`loadState`, `saveState`, `normalizeState`, `filterOperatingWeeks`) |
| 430 – 560 | MSAL 인증 (`signInMicrosoft`, `getMsalClient`, `setMicrosoftStatus`, `setMicrosoftAccountBadge`) |
| 560 – 700 | Graph API 헬퍼 (`graphFetch`, `resolveSharePointField*`, `omitSharePointFields`) |
| 700 – 920 | 멤버·의견 매퍼/저장 (`mapSharePointMember/Opinion`, `saveSharePoint*`) |
| 920 – 1700 | 나머지 리스트 매퍼/저장 (Summary, Market, Archive 등) + `loadSharePointState` |
| 1810 – 2600 | 순수 로직 (주차 유틸, 멤버 조회, 성과 계산, consensus, draft 생성) |
| 2600 – 3260 | `render()` 및 하위 `render*()` 함수 (Overview, Charts, Performance, Archive) |
| 3260 – 4020 | SVG 차트 드로잉 (`drawBase`, `drawAxes`, `drawStrategyChart` 등) |
| 4020 – 4200 | 엑셀 파싱 (`fetchSharePointRatesBlob`, `parseRatesWorkbook`, `loadBundledRatesIfAvailable`, `detectRateColumns` 등) |
| 4400 – 4980 | 이벤트 리스너 (폼 제출, 버튼 클릭) + `render()` / `loadBundledRatesIfAvailable()` 최초 호출 |

### SharePoint 리스트 (5개 운영)
| 리스트 | 용도 | 식별자/키 |
|---|---|---|
| `DAOL_FI_Members` | 구성원 원장 (11명) | `appId` = FI01~FI11 |
| `FI_WeeklyOpinions` | 의견 원장 (핵심) | `week`+`memberId`, Title=`week_memberId` |
| `FI_MarketData` | 주간 시장 데이터 | `creditAAm2y` 지뢰 |
| `FI_MarketWeekMappings` | 주차-기준일 매핑 | `baseDate` |
| `FI_WeeklySummaries` | 종합의견 | 키 = `Title` (= week) |

> Archive 7종(`FI_WeeklyArchive*` 등) + `FI_MarketMappingHistory`는 **보존만, 삭제 금지.**
> `FI_DailyRates`/`FI_UploadedRateFiles`/`FI_Settings`는 폐기 권고(D-9 일원화 후, CSV 백업 선행).

---

## 9. 거버넌스 — 일하는 방식

1. **데이터 계약(4장)은 법.** 위반 소지 있으면 멈추고 보고.
2. **검산 게이트:** 성과 계산·집계 변경은 6장 검산값과 일치해야 통과.
3. **변경 통제:** **스키마·키·부호·계산식 변경은 사용자 명시 승인 필수.** 승인 없이 바꾸지 마라.
4. **백업 선행:** 기존 데이터를 건드리는 작업 전 CSV/Excel 백업.
5. **한 번에 한 묶음:** 묶음 단위로 처리하고, 고친 뒤 **라이브 검증**(실제 저장/조회 테스트)까지가 한 작업.
6. **지시 범위 엄수:** 지시된 변경 외 리팩터링·포맷팅·정리 금지. diff는 의도한 변경만 담아야 한다.
7. **줄번호 불신:** 차터/문서의 줄번호는 사본 기준일 수 있다. **항상 문자열 검색으로 실제 위치를 재확인**하고, 형태가 예상과 다르면 멈춰 보고한다.
8. **문법 sanity:** 수정 후 구문 오류 없는지 확인하고, 변경 요약 + 커밋 메시지를 제안한다.

---

## 10. 현재 작업 상태 (transient — 진행에 따라 갱신)

**묶음 A 진행 상황:**
- ✅ **A3①** (2026-06-16) — `mapSharePointOpinion`의 `memberId` 읽기 폴백 강화: 빈 문자열도 Title(`week_memberId`)에서 역산. W26 trends 유재혁 복귀.
- ✅ **A3②** (2026-06-16) — `saveSharePointWeeklyOpinion`의 omit 목록에서 memberId 변형 4개 제거 → 저장 시 `memberId` 칸 실제 적재. 테스트 PASS.
- ✅ **A1** (2026-06-16) — 상단 바 로그인 계정 배지(`#msAccountBadge`) 추가. 작업 메시지에 덮어씌워지지 않는 전용 표시.
- ✅ **A1.1** (2026-06-16) — 로그인 상태에서 Login 버튼 숨김(`hidden` 토글). CSS `[hidden]` 강제 규칙 추가로 `.icon-button { display: inline-grid }` 충돌 해결.
- ✅ **A2** (2026-06-17) — 의견 입력 드롭다운을 로그인 본인 1명으로 고정. `mapSharePointMember`에 `email` 읽기 추가, `findSelfMemberByLogin` 헬퍼(UPN↔email 매칭), `renderMemberSelect` 본인 고정으로 교체. 잠금은 **옵션 1개 제한** 방식(`disabled` 금지 — `formData`가 disabled 값을 수집하지 않아 저장이 깨짐). 저장 가드로 빈 memberId 거부.
- ✅ **B5** (2026-06-17) — 주차 변경 시 자동조회(`loadSelectedOpinionIntoForm`). 미저장 입력 있으면 확인창으로 보호(`opinionFormHasUnsavedInput` 헬퍼).
- ✅ **A4** (2026-06-17) — 상단 바에 "최종 업데이트: YYYY-MM-DD HH:mm" (KST) 표시. `formatKstDateTime`(`Intl.DateTimeFormat` Asia/Seoul), `latestUpdatedAtIso`, `renderLastUpdatedBadge`. **묶음 A 전체 완료.**
- ✅ **B1/B2** (2026-06-17) — 2계층 하이브리드(전략 통째 skip + 개별 필드 생략) + 듀레이션 고정스텝 드롭다운(기본값 생략). R-10 재적용: `setStrategyOmitState`에서 `disabled` 제거 → `input-locked` 클래스(pointer-events:none). 레인지 생략 토글(`*RangeSkip` 체크박스) 추가. 확신도에 `생략` 옵션 추가. submit 확신도 롤업에서 null 제외(`.filter(v => v !== null)`). 비표준 레거시 듀레이션은 임시 옵션으로 사실대로 표시.
- **교훈**: 시각은 UTC(`Z`)로 저장 → 화면 표시는 항상 `Asia/Seoul`로 변환. slice로 자르면 9시간 오차.
- ✅ **B3** (2026-06-18) — `avg`/`median` null 제외 하드닝(검산 통과): `null`/`undefined`/`""`을 `Number()` 전에 필터링하여 0으로 섞이는 문제 근본 수정. trends 상세표(`durationTableRows`)에서 생략 의견을 "생략" 태그로 표시(차트·평균은 값 있는 것만). 메타에 `입력 N명` 추가. 작성자 공란(`_undefined`) 행: 빈 memberId 거부 가드로 재발 차단 확인(코드 변경 없음), 1건 복구·삭제는 사용자 SharePoint 작업.
- ✅ **B4** (2026-06-18) — 종합의견(`FI_WeeklySummaries`) 저장 경로 레거시 칸 누수 점검 완료. `rationaleText`/`rationaleTags`는 body에 없고, `confidence`는 종합의견 리스트에서는 스키마 정의된 정상 칸(전체 종합 확신도 = 전략별 평균). `writeSharePointArchiveFields`에 자동 복구 로직(invalid field 파싱→제거→재시도)도 있어 이중 안전. **누수 없음, 코드 변경 없이 종료.**
- ✅ **C2** (2026-06-18) — 금리 원천을 정적 `fetch("rates_raw.xlsx")` → SharePoint 문서 라이브러리 Graph 다운로드(`fetchSharePointRatesBlob`)로 교체. 파서·부호·차트·매핑 무변경. 미로그인 시 조용히 skip, 로그인 후 실패 시 가시적 경고. **배포 후 라이브 검증 대기.**
- **다음**: C2 라이브 검증(값 스폿체크 + 갱신 증명) → C4(`FI_DailyRates` 은퇴, CSV 백업 선행).

> 전체 백로그·결정 로그·리스크는 차터 7·9·10장 참조.

---

## 11. 핵심 파일

- `app.js` — 전 로직(인증/Graph read·write/매핑/렌더/집계). 수정 대상 대부분 여기.
- `index.html` / `styles.css` — 화면 구조·스타일.
- `vendor/msal-browser.min.js` — MSAL 라이브러리.
- `DAOL_FI_DASHBOARD_CHARTER_v2_1.md` — 프로젝트 지침서(이력·결정 근거 원본).
