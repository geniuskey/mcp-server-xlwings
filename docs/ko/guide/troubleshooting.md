# 문제 해결

## 자주 발생하는 오류

### "No active workbook found"

Excel이 실행 중이지만 워크북이 열려 있지 않거나, Excel이 실행되지 않은 상태입니다.

**해결:** Excel에서 워크북을 먼저 열거나, 다음과 같이 파일을 열어주세요:
```json
{ "action": "open", "filepath": "C:\\path\\to\\file.xlsx" }
```

### "Workbook 'xxx' is not open and the path does not exist"

지정한 워크북 이름이 열린 워크북과 일치하지 않고, 경로도 존재하지 않습니다.

**해결:** `manage_workbooks(action="list")`로 열린 워크북을 확인하거나, 전체 절대 경로를 제공하세요.

### "Sheet 'xxx' not found"

시트 이름이 일치하지 않습니다. 시트 이름은 **대소문자를 구분**합니다.

**해결:** `manage_sheets(action="list")`로 정확한 시트 이름을 확인하세요.

### "Excel COM error"

Excel 프로세스가 충돌했거나 응답하지 않거나, 대화 상자가 차단하고 있습니다.

**해결:**
1. Excel에 열린 대화 상자가 있는지 확인 (저장 프롬프트, 오류 알림 등)
2. 대화 상자를 닫고 재시도
3. Excel이 멈춘 경우 재시작

## MCP 연결 문제

### Claude Desktop: "Disconnected"

1. `claude_desktop_config.json`이 유효한 JSON인지 확인 (후행 쉼표 없음)
2. `uvx`가 시스템 PATH에 있는지 확인
3. 설정 변경 후 Claude Desktop을 완전히 재시작
4. 로그 확인: `%APPDATA%\Claude\logs\`

### Cursor / Windsurf: 서버 시작 안 됨

1. 설정 파일 경로가 정확한지 확인
2. Python 3.10+ 설치 여부 확인
3. 터미널에서 수동 테스트: `uvx mcp-server-xlwings`
4. 터미널의 오류 출력 확인

### Claude Code: 도구를 찾을 수 없음

1. `claude mcp list`로 확인
2. 없으면 재등록: `claude mcp add xlwings -- uvx mcp-server-xlwings`

## 제약 사항

| 제약 사항 | 상세 |
|----------|------|
| **Windows 전용** | COM 자동화는 Windows OS 필수 |
| **Excel 필수** | Microsoft Excel이 설치되고 라이선스가 있어야 함 |
| **헤드리스 불가** | Excel이 실행 중이어야 함 (최소화 가능, 서비스로는 불가) |
| **단일 인스턴스** | 활성 Excel 애플리케이션 인스턴스에 연결 |
| **동시 접근** | 여러 MCP 클라이언트가 동시 연결 시 충돌 가능 |
| **파일 크기** | 매우 큰 파일(100MB+)은 Excel 자체 메모리 사용으로 느릴 수 있음 |

## 자주 묻는 질문

### macOS나 Linux에서 사용할 수 있나요?

아니요. Excel COM 자동화는 Windows 전용입니다. 크로스 플랫폼이 필요하면 openpyxl 기반 MCP 서버를 고려하세요.

### Excel Online이나 Google Sheets에서 작동하나요?

아니요. Windows에서 로컬로 실행되는 데스크톱 Excel 애플리케이션이 필요합니다.

### 비밀번호로 보호된 파일을 열 수 있나요?

네, Excel이 열 수 있다면 가능합니다. 비밀번호 프롬프트가 있는 파일은 Excel에서 먼저 수동으로 여세요.

### .xls (레거시 형식)도 지원하나요?

네. Excel이 열 수 있는 모든 형식이 지원됩니다: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.csv`, `.tsv`.

### Excel에서 작업하면서 동시에 사용할 수 있나요?

네. MCP 서버와 사용자는 동일한 라이브 Excel 인스턴스를 공유합니다. 에이전트가 데이터를 읽거나 쓰는 동안에도 편집을 계속할 수 있습니다.

### 최신 버전으로 업데이트하려면?

```bash
uvx mcp-server-xlwings
```

`uvx`는 매번 PyPI에서 최신 버전을 자동으로 가져옵니다.
