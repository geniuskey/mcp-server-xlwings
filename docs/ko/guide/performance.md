# 성능 가이드

## 동작 방식

각 도구 호출은 **MCP 프로토콜 → Python → COM → Excel → 응답** 순서로 진행됩니다. 개별 COM 호출은 10~50ms가 소요되므로, 도구 호출당 COM 왕복 횟수를 최소화하는 것이 핵심입니다.

## 모범 사례

### 1. 시트 요약부터 시작

항상 `cell_range` 없이 `read_data()`를 먼저 호출하세요:

```json
{ "sheet": "Sheet1" }
```

시트 크기와 관계없이 약 5번의 COM 호출로 크기, 헤더, 병합 셀, 영역 정보를 반환합니다. 이 요약을 기반으로 읽을 범위를 결정하세요.

### 2. 배치 작업 활용

| 이렇게 하지 마세요 | 이렇게 하세요 | 절약 효과 |
|-------------------|-------------|----------|
| 시트당 5번 `read_data()` | 1번 `read_data(sheet="*")` | 5회 → 1회 |
| 셀마다 `read_data(detail=true)` | 1번 `get_formulas(cell_range="A1:Z100")` | N회 → 1회 |
| 셀마다 서식 확인 | 1번 `get_cell_styles(cell_range="...")` | N회 → 1회 |

### 3. 스타일 속성 필터링

`get_cell_styles`는 셀당 여러 COM 호출을 수행합니다. `properties`로 필요한 속성만 조회하세요:

```json
{ "cell_range": "A1:Z50", "properties": ["bold", "bg_color"] }
```

font_name, font_size, alignment, border 등을 건너뛰어 COM 호출을 크게 줄입니다.

### 4. merge_info 선택적 사용

`merge_info=true`는 범위 내 모든 셀의 병합 상태를 확인합니다 (셀당 약 1번의 COM 호출). 대규모 범위(1000셀 이상)에서는 느릴 수 있습니다. 전체 시트가 아닌 **특정 영역**에만 사용하세요.

### 5. 대용량 파일 전략

수천 행이 있는 워크북의 경우:

1. `read_data()` → 시트 요약 (크기, 헤더) 확인
2. `read_data(cell_range="A1:Z20")` → 헤더 영역을 읽어 구조 파악
3. 구조를 기반으로 대상 범위 결정
4. 특정 범위만 읽기 — 전체 시트 읽기 지양

## 성능 참고표

| 작업 | 소요 시간 | COM 호출 수 |
|------|:---------:|:----------:|
| 시트 요약 (`read_data()`) | ~50ms | ~5 |
| 100행 읽기 (벌크) | ~30ms | 1 |
| 1,000행 읽기 (벌크) | ~100ms | 1 |
| `merge_info` 100셀 | ~200ms | ~100 |
| `get_formulas` 1,000셀 | ~50ms | 2 |
| `get_cell_styles` 100셀 | ~500ms | ~100 × N 속성 |
| `get_objects` | ~20ms | ~3 |
