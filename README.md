# LG Homestyle 상품정보 파싱·규격 정규화

LG Homestyle PDP의 API, HTML, FAQ/Q&A, 상세 이미지 OCR을 결합해 상품 요구 필드와 3D Asset용 대표 규격을 생성하는 프로젝트입니다.

## 현재 데이터 기준

- 대상 상품: 9,358개
- 규격 완료: 5,713개
- 비교정보 제공: 3,411개
- 최종 무후보: 234개
- 규격 완료 상태: `SOURCE_CONFIRMED` 4,132개, `RULE_RESOLVED` 1,560개, `MANUAL_CONFIRMED` 21개
- W/D/H 중 하나라도 20mm 이하이면 의심값으로 재검증하며, 박형 제품은 명시적 제품 규격 근거가 있을 때만 예외 처리합니다.

## 주요 문서

- [개발자 인수인계 — 코드와 파이프라인](개발자_인수인계_01_코드완성도_및_파이프라인.md)
- [개발자 인수인계 — 필드 산출 데이터사전](개발자_인수인계_02_필드산출_데이터사전.md)
- [현재까지 정보 구성 Depth와 장치](현재까지_정보구성_Depth_장치_정리.md)
- [규격 표기 타입 정의](규격_표기_타입_정의_2026-07-22.md)
- [규격 문맥 정규화 최종 결과](규격_문맥정규화_최종결과_2026-07-23.md)
- [규격 비교정보 3,411개 확정 패턴](규격_비교정보3411_확정패턴목록_2026-07-27.md)
- [조합상품 3D 변환 패턴 목록](조합상품_3D변환_패턴목록_2026-07-27.md)
- [정규화 항목 및 다음 개발 백로그](정규화_항목_및_다음개발_백로그.md)
- [사용 API 목록](사용_API_목록_2026-07-22.md)

## 주요 실행 흐름

1. `bulk_homestyle_collect.py`: API·HTML·FAQ/Q&A 원천 수집
2. `run_dimension_scan_pass1.py`: 전체 상세 이미지에서 규격 후보 이미지 스캔
3. `run_dimension_scan_pass2.py`: 규격 후보 이미지 OCR 및 관찰값 적재
4. `build_dimension_context_normalization.py`: 규격 후보 문맥 분류·정규화
5. `build_dimension_resolution_ledger.py`: 확정/비교정보/무후보 상태 원장 생성
6. `build_homestyle_bulk_workbook.py`: 고객 요청 필드와 상태를 Excel로 산출

20mm 이하 품질 규칙은 `low_dimension_quality_policy.py`, 기존 오류값 보정은 `repair_invalid_locked_dimensions.py`, 20mm 이하 전수 재검증 보정은 `repair_low_dimension_values.py`에 있습니다.

## DB

기준 DB는 `homestyle_bulk_run/homestyle_bulk.sqlite`이며 Git LFS로 관리합니다.

저장소를 받은 뒤 다음 명령으로 DB를 내려받습니다.

```bash
git lfs pull
```

DB의 주요 상태 원장은 `fact_dimension_resolution_ledger`, 20mm 이하 감사 이력은 `stg_dimension_low_value_audit`입니다.

## 테스트

```bash
python -m pytest -q
```

현재 회귀 테스트는 46개입니다.

