# 아키텍처

## 개요

```
┌─────────────────┐    stdio     ┌────────────────────┐     COM      ┌─────────┐
│   MCP 클라이언트  │ ───────────→ │ mcp-server-xlwings │ ───────────→ │  Excel   │
│ (Claude, Cursor  │ ←─────────── │    (Python/MCP)    │ ←─────────── │ 프로세스  │
│  Windsurf 등)    │   JSON-RPC   └────────────────────┘  Win32 API   └─────────┘
└─────────────────┘
```

## 계층 구조

### MCP 클라이언트

Claude Desktop, Claude Code, Cursor, Windsurf, Roo Code, Continue 등 Model Context Protocol을 지원하는 모든 클라이언트입니다. stdio를 통해 JSON-RPC로 도구 호출을 전송합니다.

### MCP 서버

FastMCP로 구동되는 Python 프로세스입니다. 도구 호출을 수신하여 xlwings 작업으로 변환하고 구조화된 JSON 응답을 반환합니다. 완전히 무상태로 동작하여 세션 관리가 필요 없습니다.

### COM 자동화 (xlwings)

Windows COM을 통해 실행 중인 Excel 프로세스에 연결합니다. VBA 매크로가 사용하는 것과 동일한 인터페이스로, DRM 보호 및 암호화된 문서를 포함하여 Excel이 열어둔 모든 파일에 접근할 수 있습니다.

### Excel 프로세스

실제 Microsoft Excel 애플리케이션입니다. 파일, 수식, 매크로, 서식을 관리합니다. MCP 서버가 작업하는 동안 사용자도 동시에 Excel에서 작업할 수 있습니다.

## 설계 원칙

### 무상태

세션 ID나 연결 관리가 없습니다. 각 도구 호출이 독립적으로 활성 Excel 인스턴스를 찾습니다. 호출 사이에 Excel을 재시작해도 다음 호출이 자동으로 재연결됩니다.

### 활성 워크북 기본값

모든 도구는 `workbook`을 생략하면 활성 워크북을 기본으로 사용합니다. 가장 일반적인 경우인 사용자가 현재 열어둔 파일 작업 시 파일 경로 지정이 불필요합니다.

### 요약 우선

`read_data()`에 `cell_range`를 지정하지 않으면 전체 데이터 대신 시트 구조(크기, 헤더, 병합 셀, 영역)를 반환합니다. 에이전트가 특정 범위를 읽기 전에 시트 레이아웃을 파악할 수 있어 컨텍스트 윈도우를 절약합니다.

### 최소 도구 수

11개 도구로 Excel 작업 전체를 지원합니다. 관련 동작은 통합되어 있어 (예: `manage_workbooks`가 list/open/save/close/recalculate 처리) LLM 컨텍스트 오버헤드를 줄입니다.
