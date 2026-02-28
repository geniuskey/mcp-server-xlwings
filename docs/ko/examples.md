# 사용 예시

실제 프롬프트와 에이전트가 도구를 호출하는 과정을 카테고리별로 소개합니다. 각 예시에서 도구 호출, 파라미터, 응답 JSON을 함께 확인할 수 있습니다.

[[toc]]

---

## 시작하기

스프레드시트를 처음 열었을 때 구조를 파악하는 단계입니다.

### 스프레드시트 탐색

> "열려 있는 스프레드시트에 뭐가 있어?"

**도구 호출:**

```json
get_active_workbook()
```

**응답:**

```json
{
  "name": "매출현황_2025.xlsx",
  "path": "C:\\Users\\user\\Documents\\매출현황_2025.xlsx",
  "sheets": [
    { "name": "매출", "used_range": "$A$1:$H$120", "rows": 120, "columns": 8 },
    { "name": "비용", "used_range": "$A$1:$F$85", "rows": 85, "columns": 6 },
    { "name": "요약", "used_range": "$A$1:$D$15", "rows": 15, "columns": 4 }
  ],
  "active_sheet": "매출",
  "selection": {
    "address": "$A$1:$H$5",
    "sheet": "매출",
    "data": [["월", "지역", "제품", "수량", "단가", "매출액", "할인율", "순매출"]],
    "rows": 5,
    "columns": 8
  }
}
```

에이전트는 3개 시트의 크기와 사용 범위를 파악하고, 현재 선택 영역의 헤더를 통해 데이터 구조를 즉시 이해합니다.

### 선택한 데이터 요약

> "내가 선택한 셀 데이터를 요약해줘"

**도구 호출:**

```json
get_active_workbook()
```

**응답:**

```json
{
  "name": "매출현황_2025.xlsx",
  "selection": {
    "address": "$B$3:$E$12",
    "sheet": "매출",
    "data": [
      ["서울", "노트북", 45, 1250000],
      ["서울", "모니터", 120, 450000],
      ["부산", "노트북", 32, 1250000],
      ["부산", "모니터", 88, 450000]
    ],
    "rows": 10,
    "columns": 4
  }
}
```

`get_active_workbook()` 응답에 선택 영역 데이터가 포함되어 있으므로 추가 호출 없이 바로 분석할 수 있습니다.

### 시트 구조 분석

> "이 시트에 어떤 테이블이 있는지 알려줘"

**도구 호출:**

```json
read_data()
```

`cell_range`를 생략하면 데이터를 읽는 대신 시트의 구조 요약을 반환합니다.

**응답:**

```json
{
  "sheet": "매출",
  "used_range": "$A$1:$H$120",
  "total_rows": 120,
  "total_columns": 8,
  "headers": ["월", "지역", "제품", "수량", "단가", "매출액", "할인율", "순매출"],
  "merged_cells": [
    { "range": "$A$1:$H$1", "value": "2025년 월별 매출 현황", "rows": 1, "columns": 8 }
  ],
  "regions": [
    { "range": "$A$1:$H$60", "rows": 60, "columns": 8 },
    { "range": "$A$65:$D$80", "rows": 16, "columns": 4 }
  ]
}
```

`used_range`로 데이터가 존재하는 위치를, `merged_cells`로 병합된 제목 행을, `regions`로 분리된 테이블 경계를 파악합니다. 이 시트에는 60행짜리 메인 테이블과 16행짜리 보조 테이블이 있습니다.

---

## 읽기 및 분석

데이터를 읽고 수식, 스타일, 오브젝트 등을 분석하는 예시입니다.

### 모든 시트 일괄 읽기

> "이 통합 문서의 모든 시트를 한 번에 요약해줘"

**도구 호출:**

```json
read_data(sheet="*")
```

**응답:**

```json
{
  "sheet_count": 3,
  "sheets": {
    "매출": {
      "sheet": "매출",
      "used_range": "$A$1:$H$120",
      "total_rows": 120,
      "total_columns": 8,
      "headers": ["월", "지역", "제품", "수량", "단가", "매출액", "할인율", "순매출"]
    },
    "비용": {
      "sheet": "비용",
      "used_range": "$A$1:$F$85",
      "total_rows": 85,
      "total_columns": 6,
      "headers": ["월", "항목", "부서", "예산", "집행액", "잔액"]
    },
    "요약": {
      "sheet": "요약",
      "used_range": "$A$1:$D$15",
      "total_rows": 15,
      "total_columns": 4,
      "headers": ["구분", "금액", "비율", "비고"]
    }
  }
}
```

`sheet="*"`를 사용하면 한 번의 호출로 모든 시트의 구조를 파악할 수 있어 효율적입니다.

### 병합 셀이 있는 데이터 읽기

> "B3:D10을 읽는데 병합된 셀 값을 채워줘"

**도구 호출:**

```json
read_data(cell_range="B3:D10", merge_info=true)
```

**응답:**

```json
{
  "range": "$B$3:$D$10",
  "sheet": "매출",
  "rows": 8,
  "columns": 3,
  "headers": ["지역", "분기", "매출액"],
  "data": [
    ["서울", "1분기", 45000000],
    ["서울", "2분기", 52000000],
    ["서울", "3분기", 48000000],
    ["서울", "4분기", 61000000],
    ["부산", "1분기", 28000000],
    ["부산", "2분기", 31000000],
    ["부산", "3분기", 29000000]
  ],
  "merged_ranges": [
    { "range": "$B$4:$B$7", "value": "서울" },
    { "range": "$B$8:$B$10", "value": "부산" }
  ]
}
```

`merge_info=true`를 지정하면 병합된 셀이 null 대신 상위 셀의 값("서울", "부산")으로 채워집니다. `merged_ranges`에서 어떤 범위가 병합되어 있는지도 확인할 수 있습니다.

### 비표준 헤더 위치 처리 (header_row)

> "이 시트는 3행부터 헤더가 시작하는데, B3:G50을 읽어줘"

기업용 스프레드시트에서는 1~2행에 제목이나 빈 행이 있고, 실제 헤더가 3행 이후에 위치하는 경우가 많습니다.

**도구 호출:**

```json
read_data(cell_range="B3:G50", header_row=3)
```

**응답:**

```json
{
  "range": "$B$3:$G$50",
  "sheet": "수지분석",
  "rows": 48,
  "columns": 6,
  "headers": ["구분", "면적(m2)", "단가(원)", "금액(천원)", "비율(%)", "비고"],
  "data": [
    ["토지비", 1250.5, 3200000, 4001600, 28.5, "감정가 기준"],
    ["공사비", 1250.5, 1800000, 2250900, 16.0, "도급계약"],
    ["설계비", null, null, 180000, 1.3, "설계 용역"]
  ]
}
```

`header_row=3`을 지정하면 3행을 기준으로 헤더를 인식합니다. 1~2행에 있는 제목이나 부제목은 무시하고 실제 데이터 헤더를 정확히 매핑합니다.

### 수식 조회

> "매출 시트에 있는 모든 수식을 보여줘"

**도구 호출:**

```json
get_formulas(cell_range="A1:H120", sheet="매출", values_too=true)
```

**응답:**

```json
{
  "formulas": [
    { "cell": "F3", "formula": "=D3*E3", "value": 56250000 },
    { "cell": "H3", "formula": "=F3*(1-G3)", "value": 50625000 },
    { "cell": "F4", "formula": "=D4*E4", "value": 54000000 },
    { "cell": "H4", "formula": "=F4*(1-G4)", "value": 51300000 },
    { "cell": "F62", "formula": "=SUM(F3:F61)", "value": 1850000000 },
    { "cell": "H62", "formula": "=SUM(H3:H61)", "value": 1702500000 }
  ],
  "total_formula_cells": 236,
  "range": "$A$1:$H$120",
  "sheet": "매출"
}
```

`values_too=true`를 지정하면 수식과 계산된 값을 함께 반환합니다. F열은 `수량*단가`, H열은 `매출액*(1-할인율)`로 계산되며, 62행에 합계 수식이 있다는 것을 알 수 있습니다.

### 스타일로 헤더와 소계 식별

> "어떤 셀이 굵게 표시되거나 배경색이 있어?"

**도구 호출:**

```json
get_cell_styles(cell_range="A1:H65", sheet="매출", properties=["bold", "bg_color", "font_color"])
```

**응답:**

```json
{
  "range": "$A$1:$H$65",
  "sheet": "매출",
  "styles": [
    { "cell": "A1", "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF" },
    { "cell": "B1", "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF" },
    { "cell": "A2", "bold": true, "bg_color": "#D9E2F3" },
    { "cell": "B2", "bold": true, "bg_color": "#D9E2F3" },
    { "cell": "A62", "bold": true },
    { "cell": "F62", "bold": true },
    { "cell": "H62", "bold": true }
  ]
}
```

기본이 아닌 스타일이 적용된 셀만 반환합니다. 패턴을 분석하면:
- 진한 배경(#1B3A5C) + 흰색 글자 = 대제목 행
- 연한 배경(#D9E2F3) + 굵게 = 헤더 행
- 굵게만 적용 = 소계/합계 행

### 스타일 + 데이터 조합으로 시트 구조 파악

> "시트 구조를 스타일 기반으로 분석해줘"

복잡한 시트에서는 `get_cell_styles`와 `read_data`를 조합하면 시트의 논리적 구조를 정확히 파악할 수 있습니다.

**1단계 -- 스타일 조회:**

```json
get_cell_styles(cell_range="A1:A50", sheet="수지분석", properties=["bold", "bg_color", "font_size"])
```

```json
{
  "range": "$A$1:$A$50",
  "sheet": "수지분석",
  "styles": [
    { "cell": "A1", "bold": true, "bg_color": "#002060", "font_size": 14 },
    { "cell": "A3", "bold": true, "bg_color": "#4472C4", "font_size": 11 },
    { "cell": "A10", "bold": true, "bg_color": "#4472C4", "font_size": 11 },
    { "cell": "A20", "bold": true, "bg_color": "#4472C4", "font_size": 11 },
    { "cell": "A35", "bold": true, "bg_color": "#D6DCE4" },
    { "cell": "A45", "bold": true, "bg_color": "#D6DCE4" }
  ]
}
```

**2단계 -- 섹션 헤더 값 확인:**

```json
read_data(sheet="수지분석", cell_range="A1:A50")
```

```json
{
  "range": "$A$1:$A$50",
  "sheet": "수지분석",
  "headers": ["구분"],
  "data": [
    [null], ["수입 항목"], [null], [null], [null], [null], [null], [null],
    ["지출 항목"], [null], [null], [null], [null], [null], [null], [null], [null], [null],
    ["사업수지 요약"], [null]
  ]
}
```

스타일과 값을 조합하면 이 시트는 세 개의 주요 섹션("수입 항목", "지출 항목", "사업수지 요약")으로 구성되어 있고, 각 섹션은 파란 배경의 굵은 행으로 구분된다는 것을 알 수 있습니다.

### 차트 및 오브젝트 탐색

> "대시보드 시트에 어떤 차트가 있어?"

**도구 호출:**

```json
get_objects(sheet="대시보드")
```

**응답:**

```json
{
  "sheet": "대시보드",
  "charts": [
    {
      "name": "Chart 1",
      "top_left_cell": "$A$2",
      "width": 520,
      "height": 320,
      "chart_type": "5",
      "title": "월별 매출 추이"
    },
    {
      "name": "Chart 2",
      "top_left_cell": "$F$2",
      "width": 400,
      "height": 320,
      "chart_type": "4",
      "title": "지역별 매출 비율"
    },
    {
      "name": "Chart 3",
      "top_left_cell": "$A$20",
      "width": 520,
      "height": 280,
      "chart_type": "51",
      "title": "비용 구성"
    }
  ],
  "images": [
    { "name": "Picture 1", "top_left_cell": "$H$1", "width": 120, "height": 40 }
  ],
  "shapes": []
}
```

대시보드에 차트 3개(꺾은선형, 원형, 막대형)와 로고 이미지 1개가 있습니다. 차트의 위치(`top_left_cell`)와 크기를 통해 레이아웃을 파악할 수 있으며, `title`로 각 차트의 용도를 알 수 있습니다.

---

## 쓰기 및 서식

데이터 쓰기, 수식 추가, 서식 지정, 행 삽입 등의 예시입니다.

### 데이터 쓰기

> "A1부터 이름과 점수 데이터를 입력해줘"

**도구 호출:**

```json
write_data(
  start_cell="A1",
  data=[
    ["이름", "부서", "매출(천원)", "달성률(%)"],
    ["김민수", "영업1팀", 85000, 106.3],
    ["이지은", "영업2팀", 72000, 90.0],
    ["박서준", "영업1팀", 91000, 113.8],
    ["최유나", "영업3팀", 68000, 85.0]
  ]
)
```

**응답:**

```json
{
  "message": "Data written successfully to 매출",
  "start_cell": "A1",
  "written_range": "$A$1:$D$5",
  "rows": 5,
  "columns": 4
}
```

2D 배열로 헤더와 데이터를 한 번에 쓸 수 있습니다.

### 수식 추가

> "D6에 D2부터 D5까지의 합계 수식을 넣어줘"

**도구 호출:**

```json
write_data(start_cell="D6", formula="=SUM(D2:D5)")
```

**응답:**

```json
{
  "cell": "$D$6",
  "sheet": "매출",
  "formula": "=SUM(D2:D5)",
  "calculated_value": 316000
}
```

수식을 입력하면 Excel이 즉시 계산한 결과값을 `calculated_value`로 반환합니다.

### 서식 지정

> "헤더 행을 굵게, 가운데 정렬하고 파란 배경에 흰색 글자로 만들어줘"

**도구 호출:**

```json
format_range(
  cell_range="A1:D1",
  bold=true,
  alignment="center",
  bg_color="#1B3A5C",
  font_color="#FFFFFF",
  font_size=11,
  border=true
)
```

**응답:**

```json
{
  "range": "$A$1:$D$1",
  "sheet": "매출",
  "applied": [
    "bold=True",
    "alignment=center",
    "bg_color=#1B3A5C",
    "font_color=#FFFFFF",
    "font_size=11",
    "border=True"
  ]
}
```

여러 서식 속성을 한 번의 호출로 적용할 수 있습니다.

### 숫자 서식 적용

> "C열은 천 단위 구분자, D열은 백분율 서식으로 지정해줘"

**도구 호출 1:**

```json
format_range(cell_range="C2:C5", number_format="#,##0")
```

**도구 호출 2:**

```json
format_range(cell_range="D2:D5", number_format="0.0%")
```

**응답 1:**

```json
{
  "range": "$C$2:$C$5",
  "sheet": "매출",
  "applied": ["number_format=#,##0"]
}
```

**응답 2:**

```json
{
  "range": "$D$2:$D$5",
  "sheet": "매출",
  "applied": ["number_format=0.0%"]
}
```

`number_format`에 Excel 표준 서식 문자열을 지정합니다. 두 호출은 서로 독립적이므로 에이전트가 병렬로 실행할 수 있습니다.

### 행 삽입

> "5행에 빈 행 3개를 삽입해줘"

**도구 호출:**

```json
manage_sheets(action="insert_rows", position=5, count=3)
```

**응답:**

```json
{
  "message": "Inserted 3 row(s) at row 5.",
  "sheet": "매출"
}
```

기존 5행부터의 데이터는 아래로 밀려나고, 5~7행에 빈 행이 삽입됩니다.

---

## 검색 및 관리

찾기/바꾸기, 시트 관리, 매크로 실행, 다중 워크북 작업 예시입니다.

### 텍스트 검색

> "매출 시트에서 '서울'이 어디에 있는지 찾아줘"

**도구 호출:**

```json
find_replace(find="서울", sheet="매출")
```

**응답:**

```json
{
  "matches": [
    { "cell": "$B$3", "value": "서울", "sheet": "매출" },
    { "cell": "$B$4", "value": "서울", "sheet": "매출" },
    { "cell": "$B$5", "value": "서울 강남", "sheet": "매출" },
    { "cell": "$B$18", "value": "서울", "sheet": "매출" }
  ],
  "count": 4
}
```

`replace`를 생략하면 검색만 수행합니다. "서울 강남"처럼 부분 일치하는 셀도 포함됩니다.

### 찾기 및 바꾸기

> "모든 '2024'를 '2025'로 바꿔줘"

**도구 호출:**

```json
find_replace(find="2024", replace="2025")
```

**응답:**

```json
{
  "replaced": 47,
  "sheet": "매출"
}
```

47개의 셀에서 텍스트가 교체되었습니다.

### 시트 추가

> "'1분기 보고서'라는 새 시트를 만들어줘"

**도구 호출:**

```json
manage_sheets(action="add", new_name="1분기 보고서")
```

**응답:**

```json
{
  "message": "Sheet '1분기 보고서' added.",
  "sheets": ["매출", "비용", "요약", "1분기 보고서"]
}
```

### 시트 이름 변경

> "Sheet1 이름을 '원본데이터'로 바꿔줘"

**도구 호출:**

```json
manage_sheets(action="rename", sheet="Sheet1", new_name="원본데이터")
```

**응답:**

```json
{
  "message": "Sheet 'Sheet1' renamed to '원본데이터'.",
  "sheets": ["원본데이터", "비용", "요약"]
}
```

### VBA 매크로 실행

> "월말정산 매크로를 실행해줘"

**도구 호출:**

```json
run_macro(macro_name="월말정산", args=["2025", "03"])
```

**응답:**

```json
{
  "message": "Macro '월말정산' executed successfully.",
  "return_value": "2025년 03월 정산 완료 (처리 건수: 248)"
}
```

`args`로 매크로에 인수를 전달할 수 있습니다. 이 매크로는 연도와 월을 받아 정산을 실행합니다.

### 다중 워크북 데이터 복사

> "report.xlsx의 데이터를 summary.xlsx로 복사해줘"

**1단계 -- 원본 데이터 읽기:**

```json
read_data(workbook="report.xlsx", sheet="매출", cell_range="A1:D50")
```

```json
{
  "range": "$A$1:$D$50",
  "sheet": "매출",
  "rows": 50,
  "columns": 4,
  "headers": ["월", "제품", "수량", "매출액"],
  "data": [
    ["1월", "노트북", 120, 150000000],
    ["1월", "모니터", 85, 38250000]
  ]
}
```

**2단계 -- 대상에 쓰기:**

```json
write_data(
  workbook="summary.xlsx",
  sheet="집계",
  start_cell="A1",
  data=[["월", "제품", "수량", "매출액"], ["1월", "노트북", 120, 150000000], ["1월", "모니터", 85, 38250000]]
)
```

```json
{
  "message": "Data written successfully to 집계",
  "start_cell": "A1",
  "written_range": "$A$1:$D$3",
  "rows": 3,
  "columns": 4
}
```

에이전트가 원본에서 읽은 데이터를 `data` 파라미터에 그대로 전달하여 대상 워크북에 씁니다.

---

## 실전 워크플로우

여러 도구를 조합하여 복잡한 실무 과제를 해결하는 워크플로우입니다.

### 사업수지 분석

> "이 사업수지 분석 엑셀 파일에서 주요 재무 지표를 추출하고 정리해줘"

5개 시트, 병합 헤더, 수식이 포함된 복잡한 부동산 개발 사업수지 분석 파일을 단계별로 처리합니다.

---

**1단계 -- 워크북 구조 파악**

```json
get_active_workbook()
```

```json
{
  "name": "OO지구_사업수지분석_v3.xlsx",
  "path": "C:\\Users\\user\\Projects\\OO지구_사업수지분석_v3.xlsx",
  "sheets": [
    { "name": "사업개요", "used_range": "$A$1:$G$25", "rows": 25, "columns": 7 },
    { "name": "분양수입", "used_range": "$A$1:$N$45", "rows": 45, "columns": 14 },
    { "name": "사업비", "used_range": "$A$1:$J$80", "rows": 80, "columns": 10 },
    { "name": "자금계획", "used_range": "$A$1:$R$60", "rows": 60, "columns": 18 },
    { "name": "수지분석", "used_range": "$A$1:$H$50", "rows": 50, "columns": 8 }
  ],
  "active_sheet": "수지분석",
  "selection": {
    "address": "$A$1:$H$1",
    "sheet": "수지분석",
    "data": [["OO지구 도시개발사업 수지분석표"]],
    "rows": 1,
    "columns": 8
  }
}
```

5개 시트의 크기와 용도를 파악합니다. "사업개요"는 소규모(25행), "사업비"는 대규모(80행)입니다. 선택 영역에서 이 파일이 도시개발사업 수지분석표임을 확인합니다.

---

**2단계 -- 전체 시트 일괄 요약**

```json
read_data(sheet="*")
```

```json
{
  "sheet_count": 5,
  "sheets": {
    "사업개요": {
      "sheet": "사업개요",
      "used_range": "$A$1:$G$25",
      "total_rows": 25,
      "total_columns": 7,
      "headers": ["구분", "내용", "단위", "수량", "비고", null, null],
      "merged_cells": [
        { "range": "$A$1:$G$1", "value": "사업개요", "rows": 1, "columns": 7 }
      ],
      "regions": [
        { "range": "$A$1:$G$25", "rows": 25, "columns": 7 }
      ]
    },
    "분양수입": {
      "sheet": "분양수입",
      "used_range": "$A$1:$N$45",
      "total_rows": 45,
      "total_columns": 14,
      "headers": ["구분", "용도", "동", "층", "호수", "전용면적", "공용면적", "공급면적", "분양단가", "분양금액", "부가세", "합계", null, null],
      "merged_cells": [
        { "range": "$A$1:$N$1", "value": "분양수입 산출내역", "rows": 1, "columns": 14 },
        { "range": "$A$4:$A$15", "value": "아파트", "rows": 12, "columns": 1 },
        { "range": "$A$16:$A$28", "value": "상가", "rows": 13, "columns": 1 }
      ],
      "regions": [
        { "range": "$A$1:$N$45", "rows": 45, "columns": 14 }
      ]
    },
    "사업비": {
      "sheet": "사업비",
      "used_range": "$A$1:$J$80",
      "total_rows": 80,
      "total_columns": 10,
      "headers": ["구분", "세부항목", "산출근거", "단위", "수량", "단가(원)", "금액(천원)", "비율(%)", "비고", null],
      "merged_cells": [
        { "range": "$A$1:$J$1", "value": "총 사업비 산출내역", "rows": 1, "columns": 10 },
        { "range": "$A$4:$A$12", "value": "토지비", "rows": 9, "columns": 1 },
        { "range": "$A$13:$A$35", "value": "공사비", "rows": 23, "columns": 1 },
        { "range": "$A$36:$A$55", "value": "간접비", "rows": 20, "columns": 1 }
      ],
      "regions": [
        { "range": "$A$1:$J$80", "rows": 80, "columns": 10 }
      ]
    },
    "자금계획": {
      "sheet": "자금계획",
      "used_range": "$A$1:$R$60",
      "total_rows": 60,
      "total_columns": 18,
      "headers": ["구분", "합계", "1차", "2차", "3차", "4차", "5차", "6차", "7차", "8차"]
    },
    "수지분석": {
      "sheet": "수지분석",
      "used_range": "$A$1:$H$50",
      "total_rows": 50,
      "total_columns": 8,
      "headers": ["구분", "금액(천원)", "비율(%)", "평당금액", "비고", null, null, null],
      "merged_cells": [
        { "range": "$A$1:$H$1", "value": "OO지구 도시개발사업 수지분석표", "rows": 1, "columns": 8 }
      ],
      "regions": [
        { "range": "$A$1:$H$50", "rows": 50, "columns": 8 }
      ]
    }
  }
}
```

에이전트는 이 응답에서 다음을 파악합니다:
- "분양수입"에 아파트/상가 구분이 병합 셀로 되어 있음
- "사업비"에 토지비/공사비/간접비 세 개의 대분류가 병합 셀로 구분됨
- "수지분석"이 최종 요약 시트

---

**3단계 -- 수지분석 시트 헤더 구조 파악**

```json
read_data(sheet="수지분석", cell_range="A1:H5", merge_info=true)
```

```json
{
  "range": "$A$1:$H$5",
  "sheet": "수지분석",
  "rows": 5,
  "columns": 8,
  "headers": ["OO지구 도시개발사업 수지분석표", null, null, null, null, null, null, null],
  "data": [
    [null, null, null, null, null, null, null, null],
    ["구분", "금액(천원)", "비율(%)", "평당금액", "비고", null, null, null],
    ["I. 총수입", 85200000, 100.0, null, null, null, null, null],
    ["  1. 분양수입", 78500000, 92.1, 12800, null, null, null, null]
  ],
  "merged_ranges": [
    { "range": "$A$1:$H$1", "value": "OO지구 도시개발사업 수지분석표" }
  ]
}
```

1행은 전체 제목이 병합되어 있고, 3행이 실제 헤더입니다. 4행부터 데이터가 시작됩니다.

---

**4단계 -- 핵심 데이터 추출**

```json
read_data(sheet="수지분석", cell_range="A3:E50", header_row=3)
```

```json
{
  "range": "$A$3:$E$50",
  "sheet": "수지분석",
  "rows": 48,
  "columns": 5,
  "headers": ["구분", "금액(천원)", "비율(%)", "평당금액", "비고"],
  "data": [
    ["I. 총수입", 85200000, 100.0, null, null],
    ["  1. 분양수입", 78500000, 92.1, 12800, null],
    ["    (1) 아파트", 62000000, 72.8, 11500, null],
    ["    (2) 상가", 16500000, 19.4, 18200, null],
    ["  2. 보류지 매각", 4200000, 4.9, null, null],
    ["  3. 기타수입", 2500000, 2.9, null, "이자수입 포함"],
    ["II. 총사업비", 72800000, 85.4, null, null],
    ["  1. 토지비", 28500000, 33.4, 4650, null],
    ["    (1) 토지매입비", 25000000, 29.3, 4080, "감정평가액"],
    ["    (2) 토지부대비", 3500000, 4.1, 571, null],
    ["  2. 공사비", 31200000, 36.6, 5090, null],
    ["    (1) 직접공사비", 27800000, 32.6, 4535, "도급계약"],
    ["    (2) 간접공사비", 3400000, 4.0, 555, null],
    ["  3. 간접비", 13100000, 15.4, 2137, null],
    ["    (1) 설계비", 1800000, 2.1, 294, null],
    ["    (2) 감리비", 1200000, 1.4, 196, null],
    ["    (3) 분담금", 4500000, 5.3, 734, "기반시설"],
    ["    (4) 금융비용", 3200000, 3.8, 522, null],
    ["    (5) 제세공과금", 1400000, 1.6, 228, null],
    ["    (6) 기타경비", 1000000, 1.2, 163, null],
    ["III. 사업이익", 12400000, 14.6, 2023, null],
    ["IV. 사업이익률", null, 14.6, null, "이익/총수입"]
  ]
}
```

`header_row=3`으로 3행을 헤더로 지정하여 데이터를 정확히 매핑합니다. 에이전트는 이 데이터에서 핵심 재무 지표를 추출합니다:
- 총수입 852억원, 총사업비 728억원, 사업이익 124억원
- 사업이익률 14.6%
- 토지비 비율 33.4%, 공사비 비율 36.6%

---

**5단계 -- 수식 검증**

```json
get_formulas(sheet="수지분석", cell_range="A1:E50", values_too=true)
```

```json
{
  "formulas": [
    { "cell": "B4", "formula": "=분양수입!L30", "value": 85200000 },
    { "cell": "B5", "formula": "=분양수입!L28", "value": 78500000 },
    { "cell": "B10", "formula": "=사업비!G60", "value": 72800000 },
    { "cell": "B11", "formula": "=사업비!G12", "value": 28500000 },
    { "cell": "B14", "formula": "=사업비!G35", "value": 31200000 },
    { "cell": "B17", "formula": "=사업비!G55", "value": 13100000 },
    { "cell": "B24", "formula": "=B4-B10", "value": 12400000 },
    { "cell": "C25", "formula": "=B24/B4*100", "value": 14.6 },
    { "cell": "C5", "formula": "=B5/B4*100", "value": 92.1 },
    { "cell": "C10", "formula": "=B10/B4*100", "value": 85.4 }
  ],
  "total_formula_cells": 32,
  "range": "$A$1:$E$50",
  "sheet": "수지분석"
}
```

수식 분석 결과:
- 총수입(B4)은 "분양수입" 시트의 L30 셀에서 참조
- 총사업비(B10)는 "사업비" 시트의 G60 셀에서 참조
- 사업이익(B24) = 총수입 - 총사업비 (`=B4-B10`)
- 사업이익률(C25) = 사업이익 / 총수입 * 100 (`=B24/B4*100`)
- 각 항목의 비율은 총수입 대비 백분율로 계산됨

---

**6단계 -- 서식으로 섹션 구분 확인**

```json
get_cell_styles(sheet="수지분석", cell_range="A1:A30", properties=["bold", "bg_color", "font_size"])
```

```json
{
  "range": "$A$1:$A$30",
  "sheet": "수지분석",
  "styles": [
    { "cell": "A1", "bold": true, "bg_color": "#002060", "font_size": 14 },
    { "cell": "A3", "bold": true, "bg_color": "#D9E2F3", "font_size": 11 },
    { "cell": "A4", "bold": true, "bg_color": "#E2EFDA", "font_size": 11 },
    { "cell": "A10", "bold": true, "bg_color": "#FCE4D6", "font_size": 11 },
    { "cell": "A24", "bold": true, "bg_color": "#FFF2CC", "font_size": 12 },
    { "cell": "A25", "bold": true, "bg_color": "#FFF2CC", "font_size": 12 }
  ]
}
```

서식 패턴을 분석하면:
- 진한 파란 배경(#002060) = 문서 제목
- 연한 파란 배경(#D9E2F3) = 헤더 행
- 연한 초록 배경(#E2EFDA) = 수입 섹션
- 연한 주황 배경(#FCE4D6) = 지출 섹션
- 연한 노란 배경(#FFF2CC) = 이익/결론 섹션

에이전트는 이 6단계를 통해 사업수지를 체계적으로 분석하고, 최종적으로 다음과 같은 요약을 제공합니다:

> **사업수지 분석 요약**
> - 총수입: 852억원 (분양수입 92.1%, 보류지매각 4.9%, 기타 2.9%)
> - 총사업비: 728억원 (토지비 33.4%, 공사비 36.6%, 간접비 15.4%)
> - 사업이익: 124억원, 사업이익률 14.6%
> - 주요 비용: 토지매입비 250억(29.3%), 직접공사비 278억(32.6%), 분담금 45억(5.3%)

---

### 다중 워크북 데이터 통합

> "1월, 2월, 3월 매출 파일을 합쳐서 분기 요약 파일을 만들어줘"

3개의 월별 매출 워크북에서 데이터를 읽어 하나의 분기 요약 워크북으로 통합하는 워크플로우입니다.

---

**1단계 -- 열려 있는 워크북 확인**

```json
manage_workbooks(action="list")
```

```json
[
  { "name": "매출_2025_01.xlsx", "path": "C:\\Users\\user\\매출_2025_01.xlsx",
    "sheets": [{ "name": "매출", "used_range": "$A$1:$F$35", "rows": 35, "columns": 6 }] },
  { "name": "매출_2025_02.xlsx", "path": "C:\\Users\\user\\매출_2025_02.xlsx",
    "sheets": [{ "name": "매출", "used_range": "$A$1:$F$32", "rows": 32, "columns": 6 }] },
  { "name": "매출_2025_03.xlsx", "path": "C:\\Users\\user\\매출_2025_03.xlsx",
    "sheets": [{ "name": "매출", "used_range": "$A$1:$F$38", "rows": 38, "columns": 6 }] }
]
```

3개 파일이 모두 열려 있고, 동일한 6열 구조입니다.

---

**2단계 -- 각 워크북에서 데이터 읽기**

에이전트는 3개 워크북을 순차적으로 읽습니다.

```json
read_data(workbook="매출_2025_01.xlsx", sheet="매출", cell_range="A1:F35")
```

```json
{
  "range": "$A$1:$F$35",
  "sheet": "매출",
  "rows": 35,
  "columns": 6,
  "headers": ["날짜", "지역", "제품", "수량", "단가", "매출액"],
  "data": [
    ["2025-01-02", "서울", "노트북", 15, 1250000, 18750000],
    ["2025-01-02", "부산", "모니터", 22, 450000, 9900000],
    ["2025-01-03", "서울", "키보드", 45, 120000, 5400000]
  ]
}
```

```json
read_data(workbook="매출_2025_02.xlsx", sheet="매출", cell_range="A1:F32")
```

```json
{
  "range": "$A$1:$F$32",
  "sheet": "매출",
  "rows": 32,
  "columns": 6,
  "headers": ["날짜", "지역", "제품", "수량", "단가", "매출액"],
  "data": [
    ["2025-02-03", "서울", "노트북", 18, 1250000, 22500000],
    ["2025-02-03", "대전", "모니터", 12, 450000, 5400000]
  ]
}
```

```json
read_data(workbook="매출_2025_03.xlsx", sheet="매출", cell_range="A1:F38")
```

```json
{
  "range": "$A$1:$F$38",
  "sheet": "매출",
  "rows": 38,
  "columns": 6,
  "headers": ["날짜", "지역", "제품", "수량", "단가", "매출액"],
  "data": [
    ["2025-03-02", "부산", "노트북", 20, 1250000, 25000000],
    ["2025-03-02", "서울", "모니터", 30, 450000, 13500000]
  ]
}
```

3개 파일 모두 동일한 헤더 구조("날짜", "지역", "제품", "수량", "단가", "매출액")를 가지고 있음을 확인합니다.

---

**3단계 -- 요약 워크북 생성**

```json
manage_workbooks(action="open", filepath="new")
```

```json
{
  "name": "통합 문서1",
  "sheets": [{ "name": "Sheet1", "used_range": "", "rows": 0, "columns": 0 }]
}
```

빈 워크북이 생성되었습니다.

---

**4단계 -- 시트 이름 변경 및 구조 준비**

```json
manage_sheets(action="rename", workbook="통합 문서1", sheet="Sheet1", new_name="1분기 요약")
```

```json
{
  "message": "Sheet 'Sheet1' renamed to '1분기 요약'.",
  "sheets": ["1분기 요약"]
}
```

---

**5단계 -- 집계 데이터 쓰기**

에이전트가 3개 워크북의 데이터를 지역별/제품별로 집계하여 씁니다.

```json
write_data(
  workbook="통합 문서1",
  sheet="1분기 요약",
  start_cell="A1",
  data=[
    ["2025년 1분기 매출 요약"],
    [],
    ["지역", "제품", "1월 수량", "1월 매출", "2월 수량", "2월 매출", "3월 수량", "3월 매출", "합계 수량", "합계 매출"],
    ["서울", "노트북", 42, 52500000, 51, 63750000, 48, 60000000, 141, 176250000],
    ["서울", "모니터", 65, 29250000, 58, 26100000, 72, 32400000, 195, 87750000],
    ["서울", "키보드", 130, 15600000, 145, 17400000, 118, 14160000, 393, 47160000],
    ["부산", "노트북", 28, 35000000, 32, 40000000, 38, 47500000, 98, 122500000],
    ["부산", "모니터", 48, 21600000, 42, 18900000, 55, 24750000, 145, 65250000],
    ["대전", "노트북", 15, 18750000, 18, 22500000, 22, 27500000, 55, 68750000],
    ["대전", "모니터", 20, 9000000, 12, 5400000, 18, 8100000, 50, 22500000]
  ]
)
```

```json
{
  "message": "Data written successfully to 1분기 요약",
  "start_cell": "A1",
  "written_range": "$A$1:$J$10",
  "rows": 10,
  "columns": 10
}
```

---

**6단계 -- 합계 행 및 수식 추가**

```json
write_data(workbook="통합 문서1", sheet="1분기 요약", start_cell="A11", data=[["합계"]])
```

```json
write_data(workbook="통합 문서1", sheet="1분기 요약", start_cell="C11", formula="=SUM(C4:C10)")
```

```json
{
  "cell": "$C$11",
  "sheet": "1분기 요약",
  "formula": "=SUM(C4:C10)",
  "calculated_value": 348
}
```

나머지 열(D~J)에도 동일하게 SUM 수식을 추가합니다.

---

**7단계 -- 서식 적용**

```json
format_range(
  workbook="통합 문서1",
  sheet="1분기 요약",
  cell_range="A1:J1",
  bold=true, font_size=14, bg_color="#002060", font_color="#FFFFFF"
)
```

```json
format_range(
  workbook="통합 문서1",
  sheet="1분기 요약",
  cell_range="A3:J3",
  bold=true, alignment="center", bg_color="#D9E2F3", border=true
)
```

```json
format_range(
  workbook="통합 문서1",
  sheet="1분기 요약",
  cell_range="D4:D11",
  number_format="#,##0"
)
```

```json
format_range(
  workbook="통합 문서1",
  sheet="1분기 요약",
  cell_range="A11:J11",
  bold=true, bg_color="#FFF2CC", border=true
)
```

제목, 헤더, 금액 서식, 합계 행에 각각 서식을 적용합니다.

---

**8단계 -- 파일 저장**

```json
manage_workbooks(action="save", workbook="통합 문서1", filepath="C:\\Users\\user\\매출_2025_Q1_요약.xlsx")
```

```json
{
  "message": "Workbook saved as 'C:\\Users\\user\\매출_2025_Q1_요약.xlsx'."
}
```

3개 월별 파일의 데이터를 읽고, 지역/제품별로 집계하고, 서식을 적용하여 분기 요약 파일을 완성했습니다.
