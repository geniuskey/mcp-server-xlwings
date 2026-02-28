# 도구 레퍼런스

모든 도구는 `workbook`을 생략하면 **활성 통합 문서**를 기본으로 사용합니다.

## get_active_workbook

현재 활성 통합 문서의 정보를 가져옵니다. 시트별 메타데이터와 현재 선택 영역 및 데이터를 포함합니다.

**파라미터:** 없음

**응답:**

```json
{
  "name": "report.xlsx",
  "path": "C:\\Users\\user\\report.xlsx",
  "sheets": [
    { "name": "Data", "used_range": "$A$1:$Z$100", "rows": 100, "columns": 26 },
    { "name": "Summary", "used_range": "$A$1:$D$10", "rows": 10, "columns": 4 }
  ],
  "active_sheet": "Data",
  "selection": {
    "address": "$A$1:$C$5",
    "sheet": "Data",
    "data": [["ID", "Name", "Value"], [1, "Alice", 100], ...],
    "rows": 5,
    "columns": 3
  }
}
```

---

## manage_workbooks

Excel 통합 문서를 관리합니다: 목록 조회, 열기, 저장, 닫기, 재계산.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `action` | string | Yes | `list`, `open`, `save`, `close`, `recalculate` |
| `workbook` | string | No | 통합 문서 이름 또는 경로. 기본값은 활성 통합 문서 |
| `filepath` | string | No | `open`: 파일 경로 (빈 파일은 `"new"` 사용). `save`: 다른 이름으로 저장 경로 |
| `read_only` | bool | No | `open`: 읽기 전용으로 열기 (기본값 `true`) |
| `save` | bool | No | `close`: 닫기 전 저장 여부 |

**예시 -- 열린 통합 문서 목록 조회:**

```json
// Request
{ "action": "list" }

// Response
[
  {
    "name": "report.xlsx",
    "path": "C:\\Users\\user\\report.xlsx",
    "sheets": [
      { "name": "Sheet1", "used_range": "$A$1:$D$50", "rows": 50, "columns": 4 }
    ]
  }
]
```

---

## read_data

Excel 범위에서 데이터를 읽습니다. `cell_range`를 생략하면 전체 데이터를 읽는 대신 구조 분석이 포함된 **시트 요약**을 반환합니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 기본값은 활성 시트. 모든 시트를 읽으려면 `*`를 사용 |
| `cell_range` | string | No | `A1:D10`과 같은 범위. 생략하면 시트 요약을 반환 |
| `headers` | bool | No | 첫 번째 행을 헤더로 처리 (기본값 `true`) |
| `detail` | bool | No | 단일 셀: 수식, 타입, 서식 정보 포함 |
| `merge_info` | bool | No | 병합된 셀을 null 대신 병합 영역의 값으로 채움 (기본값 `false`) |
| `header_row` | int | No | 헤더로 사용할 1 기반 행 번호 (예: `3`은 3행을 의미) |

**예시 -- 시트 요약 (범위 없음):**

```json
// Request
{ }

// Response
{
  "sheet": "Sheet1",
  "used_range": "$B$3:$Z$100",
  "total_rows": 98,
  "total_columns": 25,
  "headers": ["2024 Revenue", null, null, "Q1", "Q2", "Q3"],
  "merged_cells": [
    { "range": "$B$3:$H$3", "value": "2024 Revenue", "rows": 1, "columns": 7 },
    { "range": "$B$4:$D$4", "value": "Q1", "rows": 1, "columns": 3 }
  ],
  "regions": [
    { "range": "$B$3:$H$20", "rows": 18, "columns": 7 },
    { "range": "$B$25:$F$40", "rows": 16, "columns": 5 }
  ]
}
```

**예시 -- 특정 범위 읽기:**

```json
// Request
{ "cell_range": "A1:D10" }

// Response
{
  "range": "$A$1:$D$10",
  "sheet": "Sheet1",
  "rows": 10,
  "columns": 4,
  "headers": ["ID", "Name", "Date", "Amount"],
  "data": [
    [1, "Alice", "2024-01-15", 1500],
    [2, "Bob", "2024-01-16", 2300]
  ]
}
```

**예시 -- 단일 셀 상세 정보:**

```json
// Request
{ "cell_range": "C10", "detail": true }

// Response
{
  "range": "$C$10",
  "sheet": "Sheet1",
  "rows": 1,
  "columns": 1,
  "data": [[3800]],
  "detail": {
    "value": 3800,
    "type": "number",
    "formula": "=SUM(C2:C9)",
    "number_format": "#,##0",
    "font": { "name": "Calibri", "size": 11, "bold": true, "italic": false }
  }
}
```

**예시 -- merge_info:**

```json
// Request
{ "cell_range": "B6:C8", "merge_info": true }

// Response
{
  "range": "$B$6:$C$8",
  "sheet": "Sheet1",
  "rows": 3,
  "columns": 2,
  "headers": ["건축 계획", "대지"],
  "data": [
    ["건축 계획", "면적"],
    ["건축 계획", null]
  ],
  "merged_ranges": [
    { "range": "$B$6:$B$19", "value": "건축 계획" }
  ]
}
```

**예시 -- 모든 시트 일괄 읽기:**

```json
// Request
{ "sheet": "*" }

// Response
{
  "sheet_count": 3,
  "sheets": {
    "Sheet1": { "sheet": "Sheet1", "used_range": "$A$1:$D$50", "total_rows": 50, "total_columns": 4, "headers": ["ID", "Name", "Date", "Amount"] },
    "Sheet2": { "sheet": "Sheet2", "used_range": "$A$1:$B$20", "total_rows": 20, "total_columns": 2, "headers": ["Category", "Total"] },
    "Summary": { "sheet": "Summary", "used_range": "$A$1:$C$5", "total_rows": 5, "total_columns": 3, "headers": ["Metric", "Value", "Change"] }
  }
}
```

---

## write_data

Excel 셀에 데이터 또는 수식을 작성합니다. 2D 배열은 `data`를, 단일 셀 수식은 `formula`를 사용합니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `start_cell` | string | Yes | 시작 셀 (예: `A1`) |
| `data` | array | No | 2D 값 목록. `formula`와 동시 사용 불가 |
| `formula` | string | No | `=SUM(A1:A10)`과 같은 Excel 수식. `data`와 동시 사용 불가 |
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 기본값은 활성 시트 |

**예시 -- 데이터 작성:**

```json
// Request
{
  "start_cell": "A1",
  "data": [["Name", "Score"], ["Alice", 95], ["Bob", 87]]
}

// Response
{
  "message": "Data written successfully to Sheet1",
  "start_cell": "A1",
  "written_range": "$A$1:$B$3",
  "rows": 3,
  "columns": 2
}
```

**예시 -- 수식 설정:**

```json
// Request
{ "start_cell": "C10", "formula": "=SUM(C2:C9)" }

// Response
{
  "cell": "$C$10",
  "sheet": "Sheet1",
  "formula": "=SUM(C2:C9)",
  "calculated_value": 3800
}
```

---

## manage_sheets

시트 및 구조를 관리합니다: 목록 조회, 추가, 삭제, 이름 변경, 복사, 활성화, 행/열 삽입 및 삭제.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `action` | string | Yes | `list`, `add`, `delete`, `rename`, `copy`, `activate`, `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns` |
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 대상 시트 (delete/rename/copy/activate에 필수) |
| `new_name` | string | No | 새 이름 (rename용; add/copy에서는 선택 사항) |
| `position` | int | No | 행/열 번호 (1 기반), 삽입/삭제용 (기본값 `1`) |
| `count` | int | No | 행/열 개수 (기본값 `1`) |

**예시 -- 행 삽입:**

```json
// Request
{ "action": "insert_rows", "position": 5, "count": 3 }

// Response
{ "message": "Inserted 3 row(s) at row 5.", "sheet": "Sheet1" }
```

---

## find_replace

시트에서 텍스트를 검색하고, 선택적으로 대체합니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `find` | string | Yes | 검색할 텍스트 |
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 기본값은 활성 시트 |
| `replace` | string | No | 대체 텍스트. 검색만 하려면 생략 |
| `match_case` | bool | No | 대소문자 구분 (기본값 `false`) |

**예시 -- 검색:**

```json
// Request
{ "find": "Alice" }

// Response
{
  "matches": [
    { "cell": "$A$2", "value": "Alice", "sheet": "Sheet1" },
    { "cell": "$A$15", "value": "Alice Smith", "sheet": "Sheet1" }
  ],
  "count": 2
}
```

---

## format_range

셀 범위에 서식을 적용합니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `cell_range` | string | Yes | `A1:D10`과 같은 범위 |
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 기본값은 활성 시트 |
| `bold` | bool | No | 굵게 설정 |
| `italic` | bool | No | 기울임꼴 설정 |
| `underline` | bool | No | 밑줄 설정 |
| `font_size` | int | No | 글꼴 크기 (포인트) |
| `font_color` | string | No | 16진수 색상 코드 (예: `#FF0000`) |
| `bg_color` | string | No | 배경 16진수 색상 코드 (예: `#FFFF00`) |
| `number_format` | string | No | Excel 서식 (예: `#,##0.00`) |
| `alignment` | string | No | `left`, `center`, `right`, `justify` |
| `wrap_text` | bool | No | 텍스트 줄 바꿈 활성화 |
| `border` | bool | No | 얇은 테두리 적용 |

**예시:**

```json
// Request
{ "cell_range": "A1:D1", "bold": true, "alignment": "center", "bg_color": "#FFFF00" }

// Response
{
  "range": "$A$1:$D$1",
  "sheet": "Sheet1",
  "applied": ["bold=True", "alignment=center", "bg_color=#FFFF00"]
}
```

---

## run_macro

Excel에서 VBA 매크로를 실행하고 결과를 반환합니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `macro_name` | string | Yes | 매크로 이름 (예: `MyMacro` 또는 `Module1.MyMacro`) |
| `workbook` | string | No | 통합 문서 이름. 생략하면 Excel이 전역으로 검색 |
| `args` | array | No | 매크로에 전달할 인수 |

**예시:**

```json
// Request
{ "macro_name": "UpdateReport" }

// Response
{
  "message": "Macro 'UpdateReport' executed successfully.",
  "return_value": "Report updated"
}
```

---

## get_formulas

범위 내 모든 수식을 가져옵니다. 수식이 포함된 셀만 반환합니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `cell_range` | string | Yes | `A1:U99`와 같은 범위 |
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 기본값은 활성 시트 |
| `values_too` | bool | No | 수식과 함께 계산된 값도 포함 (기본값 `false`) |

**예시:**

```json
// Request
{ "cell_range": "A1:U99", "values_too": true }

// Response
{
  "formulas": [
    { "cell": "J22", "formula": "=E22*H22", "value": 37828200 },
    { "cell": "K22", "formula": "=J22/$J$36", "value": 0.3739 }
  ],
  "total_formula_cells": 2,
  "range": "$A$1:$U$99",
  "sheet": "Sheet1"
}
```

---

## get_cell_styles

범위 내 셀의 서식/스타일 정보를 가져옵니다. 기본 스타일이 아닌 셀만 반환합니다. 시각적 서식을 통해 헤더, 소계, 데이터 역할을 식별하는 데 유용합니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `cell_range` | string | Yes | `A1:D10`과 같은 범위 |
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 기본값은 활성 시트 |
| `properties` | array | No | 특정 속성 필터: `bold`, `italic`, `underline`, `font_name`, `font_size`, `font_color`, `bg_color`, `number_format`, `alignment`, `border` |

**예시:**

```json
// Request
{ "cell_range": "A1:D5", "properties": ["bold", "bg_color", "font_color"] }

// Response
{
  "range": "$A$1:$D$5",
  "sheet": "Sheet1",
  "styles": [
    { "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF", "cell": "A1" },
    { "bold": true, "bg_color": "#1B3A5C", "font_color": "#FFFFFF", "cell": "B1" },
    { "bold": true, "cell": "A5" }
  ]
}
```

---

## get_objects

시트의 차트, 이미지, 도형 목록을 가져옵니다.

**파라미터:**

| 파라미터 | 타입 | 필수 | 설명 |
|-----------|------|----------|-------------|
| `workbook` | string | No | 기본값은 활성 통합 문서 |
| `sheet` | string | No | 기본값은 활성 시트 |

**예시:**

```json
// Request
{ }

// Response
{
  "sheet": "Dashboard",
  "charts": [
    {
      "name": "Chart 1",
      "top_left_cell": "$F$2",
      "width": 480,
      "height": 300,
      "chart_type": "5",
      "title": "Monthly Revenue"
    }
  ],
  "images": [
    { "name": "Picture 1", "top_left_cell": "$A$1", "width": 200, "height": 80 }
  ],
  "shapes": []
}
```
