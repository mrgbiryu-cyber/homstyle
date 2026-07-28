# DB 필수값 판정 스테이징 정의

기준일: 2026-07-22

## 목적

엑셀을 만들기 전에 상품별 필수값 PASS/보강대상 결과를 SQLite에서 먼저
고정하고 검증한다. 엑셀 통계와 상품별 판정은 이 스테이징 결과를 사용한다.

SQLite `TEMP TABLE`은 연결 종료 시 사라지므로, 재현성과 비교를 위해 일반
영구 테이블로 만들고 스냅샷 방식으로 관리한다.

## DB와 테이블

- DB: `homestyle_bulk_run/homestyle_bulk.sqlite`
- 테이블: `stg_mandatory_pass`
- 기본키: `(snapshot_id, product_id)`
- 현재 스냅샷: `is_current=1`
- 과거 스냅샷: 새 실행 시 `is_current=0`으로 보존

## 주요 컬럼

| 컬럼 | 설명 |
|---|---|
| `snapshot_id` | 판정 실행 스냅샷 ID |
| `assessed_at` | 판정 시각 |
| `standard_source` | 필수값 기준 원천 PPTX |
| `standard_version` | `0715_ing_R` |
| `parser_version` | 적용한 규격·필수값 판정 버전 |
| `product_id` | 상품 고유키 |
| `final_status` | `PASS` 또는 `보강대상` |
| `fulfilled_count` | 충족한 필수 그룹 수 |
| `required_count` | 해당 상품에 적용되는 필수 그룹 수 |
| `missing_fields` | 미충족 필수 그룹 목록 |
| `collection_error` | 상품 수집오류 여부 |
| `image_requirement_status` | 현재 `URL 대체충족` 또는 `미확보` |
| `is_set_applicable` | 세트상품 여부(세트 구성 ID 참고 필드 적용 여부) |
| `*_ok` | 필수 그룹 충족=1, 미충족=0, 비적용=NULL |
| `brand_value` 등 | 판정에 사용한 실제 원천값 |
| `w_mm`, `d_mm`, `h_mm` | PASS 판정에 사용한 숫자 규격 |
| `rolling_image_urls_json` | 판정에 사용한 롤링 이미지 URL 목록 |

## 통계 조회 뷰

### 현재 PASS/보강대상

```sql
SELECT *
FROM vw_mandatory_pass_current_summary;
```

### 현재 필수 그룹별 누락

```sql
SELECT *
FROM vw_mandatory_pass_current_missing
WHERE product_count > 0
ORDER BY product_count DESC;
```

### 보강대상 상품 목록

```sql
SELECT product_id, product_name, fulfilled_count, required_count, missing_fields
FROM stg_mandatory_pass
WHERE is_current=1 AND final_status='보강대상'
ORDER BY product_id;
```

### 특정 필수값 보강대상

```sql
SELECT product_id, product_name, w_mm, d_mm, h_mm
FROM stg_mandatory_pass
WHERE is_current=1 AND size_wdh_ok=0
ORDER BY product_id;
```

## 현재 스냅샷

- 스냅샷 ID: `2026-07-22T145518_0900`
- 상품 행: 9,358
- PASS: 3,955
- 보강대상: 5,403
- PASS율: 42.3%
- 사이즈 W/D/H 보강: 5,208
- 세트 구성 실제 ID: 참고 통계로만 유지하며 필수 보강 통계에서는 제외(2026-07-22)
- 색상 보강: 737

보강 필드 건수는 한 상품에 여러 누락이 있으면 중복 집계된다.

## 보강 원인·실행 백로그

`stg_reinforcement_backlog`은 현재 보강대상 5,403개를 상품×누락 필드
5,945행으로 분리한다. 색상과 규격이 함께 누락된 상품도 두 작업으로 각각
추적한다.

주요 컬럼:

| 컬럼 | 의미 |
|---|---|
| `missing_field` | `색상` 또는 `규격(W/D/H)` |
| `missing_combination` | 상품에 동시에 누락된 필드 조합 |
| `cause_code`, `cause_label` | 실제 미확보 원인 코드와 설명 |
| `candidate_value`, `candidate_evidence` | 현재 원천에서 찾은 후보와 근거(아직 PASS에 미적용) |
| `proposed_method` | 다음 보강 실행 방식 |
| `automation_level` | 자동 반영·검수·OCR·시각 추론·정책·수기 등 실행 구분 |
| `priority` | 낮을수록 먼저 처리할 우선순위 |
| `details_json` | 옵션/OCR/이미지/누락 축 등 진단 상세 |

원인별 집계:

```sql
SELECT *
FROM vw_reinforcement_backlog_current_summary
ORDER BY missing_field, product_count DESC;
```

상품별 누락 작업:

```sql
SELECT *
FROM vw_reinforcement_backlog_current_products
ORDER BY first_priority, product_id;
```

갱신 프로그램:

```powershell
python build_reinforcement_backlog.py
```

결과 JSON: `homestyle_bulk_run/reinforcement_backlog_latest.json`

사람이 읽는 분류표: `보강대상_분류_및_보강방법_2026-07-22.md`

## 규격 표기 타입 스테이징

`요청1_대표 규격 대상`은 제품 외형·옵션·구성품 구분으로 유지하고,
원문의 W/D/H·L/D/H 등 표기 패턴은 별도 타입으로 관리한다.

DB 객체:

- `ref_dimension_notation_type`: 규격 표기 타입 기준표
- `stg_dimension_pattern`: 상품별 대표 대상 타입, 표기 타입, 축 순서, 매핑 상태
- `vw_dimension_pattern_current_summary`: 타입별 현재 통계
- `vw_dimension_reinforcement_with_pattern`: 규격 보강대상과 타입의 결합 결과

갱신 프로그램:

```powershell
python build_dimension_pattern_staging.py
```

상세 정의: `규격_표기_타입_정의_2026-07-22.md`

## 갱신 프로그램

```powershell
python build_mandatory_pass_staging.py
```

실행 순서:

1. 현재 API·HTML·FAQ/Q&A·OCR 및 규격 정규화 결과를 읽는다.
2. 별첨 0715 기준으로 상품별 필수 그룹을 판정한다.
3. 기존 현재 스냅샷을 과거 상태로 보존한다.
4. 새 스냅샷을 `is_current=1`로 삽입한다.
5. PASS/보강대상 및 누락 필드 통계를 JSON으로 함께 저장한다.

결과 JSON: `homestyle_bulk_run/mandatory_staging_latest.json`

## Excel 연계 원칙

- Excel 생성 전에 스테이징의 현재 행 수가 대상 상품 수와 같은지 확인한다.
- `00_요약`의 PASS/보강대상 수는 현재 스냅샷에서 집계한다.
- `01_상품별_요구필드`의 필수값 판정은 `product_id`로 결합한다.
- 보강대상 상세는 현재 스냅샷의 `missing_fields`를 출력한다.
- 스테이징과 실시간 재계산 값이 다르면 엑셀 생성을 중단한다.
