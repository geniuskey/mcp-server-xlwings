# 시작하기

## 요구 사항

- **Windows** (Excel COM 자동화는 Windows에서만 지원됩니다)
- **Microsoft Excel** 설치 필요
- **Python 3.10+**

## 설치

### uvx 사용 (권장)

별도 설치가 필요 없습니다. MCP 클라이언트가 자동으로 실행합니다:

```bash
uvx mcp-server-xlwings
```

### pip 사용

```bash
pip install mcp-server-xlwings
```

## 빠른 시작

1. 서버를 설치합니다 (또는 `uvx`를 사용합니다)
2. MCP 클라이언트에 설정을 추가합니다 ([설정](./configuration) 참고)
3. Excel 파일을 엽니다
4. AI 에이전트에게 스프레드시트를 읽거나 수정하도록 요청합니다

서버는 COM 자동화를 통해 실행 중인 Excel 프로세스에 연결됩니다. 도구가 작동하려면 **Excel이 실행 중이어야 합니다**.

## 동작 원리

```
AI 에이전트  <-->  MCP 클라이언트  <-->  mcp-server-xlwings  <-->  Excel (COM)
```

파일 기반 라이브러리(openpyxl, pandas)와 달리, 이 서버는 **실행 중인 Excel 프로세스**와 직접 통신합니다. 이를 통해 다음과 같은 장점이 있습니다:

- DRM으로 보호된 파일도 원활하게 작동합니다
- 수식의 계산된 값을 반환합니다
- VBA 매크로를 실행할 수 있습니다
- 사용자의 현재 선택 영역에 접근할 수 있습니다
- 변경 사항이 Excel 창에 실시간으로 반영됩니다
