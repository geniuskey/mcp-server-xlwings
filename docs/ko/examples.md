# 사용 예시

실제 프롬프트와 에이전트가 도구를 사용하는 방법을 소개합니다.

## 스프레드시트 탐색

> "열려 있는 스프레드시트에 뭐가 있어?"

1. 에이전트가 `get_active_workbook()`을 호출합니다 -- 통합 문서 이름, 시트별 메타데이터(행, 열, 사용 범위), 현재 선택 영역 데이터를 가져옵니다
2. 에이전트가 `read_data()`를 호출합니다 -- 헤더, 병합 셀, 영역이 포함된 시트 요약을 가져옵니다
3. 에이전트가 `read_data(cell_range="A1:D20")`를 호출합니다 -- 특정 데이터 범위를 가져옵니다

## 선택한 데이터 요약

> "내가 선택한 데이터를 요약해줘"

에이전트가 `get_active_workbook()`을 호출합니다 -- 응답에 선택 영역 데이터가 직접 포함되어 있어 추가 호출이 필요 없습니다.

## 복잡한 시트 분석

> "이 시트에 어떤 테이블이 있어?"

에이전트가 범위 없이 `read_data()`를 호출합니다. 응답에 다음이 포함됩니다:
- `used_range`: 데이터가 존재하는 위치
- `merged_cells`: 계층적 헤더 구조
- `regions`: 구분된 테이블 경계

```json
{
  "used_range": "$B$3:$Z$100",
  "merged_cells": [
    { "range": "$B$3:$H$3", "value": "2024 Revenue", "rows": 1, "columns": 7 }
  ],
  "regions": [
    { "range": "$B$3:$H$20", "rows": 18, "columns": 7 },
    { "range": "$B$25:$F$40", "rows": 16, "columns": 5 }
  ]
}
```

## 모든 시트 한 번에 읽기

> "이 통합 문서의 모든 시트를 요약해줘"

에이전트가 `read_data(sheet="*")`를 호출합니다 -- 한 번의 호출로 모든 시트의 요약을 반환합니다.

## 병합 셀이 있는 데이터 읽기

> "B6:C20을 읽고 병합된 셀 값을 채워줘"

에이전트가 `read_data(cell_range="B6:C20", merge_info=true)`를 호출합니다. 병합된 셀이 null 대신 상위 셀의 값으로 채워집니다.

## 모든 수식 찾기

> "매출 시트에 있는 모든 수식을 보여줘"

에이전트가 `get_formulas(cell_range="A1:Z100", sheet="Revenue", values_too=true)`를 호출합니다 -- 모든 수식과 계산된 값을 반환합니다.

## 스타일로 헤더 식별

> "어떤 셀이 굵게 표시되거나 배경색이 있어?"

에이전트가 `get_cell_styles(cell_range="A1:Z10", properties=["bold", "bg_color"])`를 호출합니다 -- 기본이 아닌 굵게 또는 배경색이 있는 셀만 반환합니다.

## 차트 및 이미지 확인

> "이 시트에 차트나 이미지가 있어?"

에이전트가 `get_objects()`를 호출합니다 -- 모든 차트, 이미지, 도형의 위치와 크기를 나열합니다.

## 합계 행 만들기

> "C10에 C2:C9를 합산하는 SUM 수식을 추가해줘"

에이전트가 `write_data(start_cell="C10", formula="=SUM(C2:C9)")`를 호출하고 계산된 값을 즉시 받습니다.

## 헤더 행 서식 지정

> "1행을 굵게, 가운데 정렬하고 노란색 배경으로 만들어줘"

에이전트가 `format_range(cell_range="A1:D1", bold=true, alignment="center", bg_color="#FFFF00")`를 호출합니다.

## 행 삽입

> "5행에 빈 행 3개를 삽입해줘"

에이전트가 `manage_sheets(action="insert_rows", position=5, count=3)`를 호출합니다.

## VBA 매크로 실행

> "UpdateReport 매크로를 실행해줘"

에이전트가 `run_macro(macro_name="UpdateReport")`를 호출하고 결과를 반환합니다.

## 여러 통합 문서로 작업

> "report.xlsx의 데이터를 summary.xlsx로 복사해줘"

1. 에이전트가 `read_data(workbook="report.xlsx", cell_range="A1:D50")`를 호출합니다
2. 에이전트가 `write_data(workbook="summary.xlsx", start_cell="A1", data=...)`를 호출합니다

## 찾기 및 바꾸기

> "모든 '2024'를 '2025'로 바꿔줘"

에이전트가 `find_replace(find="2024", replace="2025")`를 호출합니다.

## 시트 관리

> "'Q1 Report'라는 새 시트를 만들어줘"

에이전트가 `manage_sheets(action="add", new_name="Q1 Report")`를 호출합니다.
