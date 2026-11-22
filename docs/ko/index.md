---
layout: home
title: mcp-server-xlwings — AI 에이전트를 위한 Excel 자동화
titleTemplate: false

hero:
  name: mcp-server-xlwings
  text: AI 에이전트를 위한 Excel 자동화
  tagline: COM을 통해 실행 중인 Excel을 직접 제어하는 MCP 서버. DRM 보호 파일도 지원합니다. Windows 전용.
  image:
    src: /logo.svg
    alt: mcp-server-xlwings
  actions:
    - theme: brand
      text: 시작하기
      link: /ko/guide/getting-started
    - theme: alt
      text: GitHub
      link: https://github.com/geniuskey/mcp-server-xlwings

features:
  - icon: 🔒
    title: DRM 보호 파일 지원
    details: COM 자동화로 실행 중인 Excel 프로세스와 직접 통신합니다. Excel이 열 수 있는 모든 파일을 읽고 쓸 수 있습니다.
  - icon: ⚡
    title: 실시간 Excel 제어
    details: 선택 영역 읽기, VBA 매크로 실행, 수식 결과 즉시 확인, 재계산 등 파일 기반 라이브러리로는 불가능한 기능을 제공합니다.
  - icon: 🛠️
    title: 11개 도구
    details: 읽기, 쓰기, 서식, 검색, 수식 조회, 스타일 조회, 차트 감지, 매크로 실행까지 통합된 도구 세트.
  - icon: 📊
    title: 스마트 시트 분석
    details: 병합 셀 자동 감지, 데이터 영역 탐색, 시트 구조 분석으로 복잡한 기업용 스프레드시트도 파악할 수 있습니다.
  - icon: 🔄
    title: 배치 작업
    details: sheet="*"로 전체 시트 일괄 읽기, 수식 일괄 조회, 범위 내 셀 스타일 한 번에 조회가 가능합니다.
  - icon: 🚀
    title: 설정 없이 바로 사용
    details: uvx 한 줄이면 설치 끝. Claude Desktop, Claude Code, Cursor, Windsurf, Roo Code, Continue를 지원합니다.
---

## 빠른 시작

MCP 클라이언트 설정에 추가하면 바로 사용할 수 있습니다:

```json
{
  "mcpServers": {
    "xlwings": {
      "command": "uvx",
      "args": ["mcp-server-xlwings"]
    }
  }
}
```

AI 에이전트에게 물어보세요:

> "지금 열려있는 엑셀 파일 내용이 뭐야?"

에이전트가 `get_active_workbook()`과 `read_data()`를 호출하여 Excel 파일을 분석합니다.

[시작하기 →](/ko/guide/getting-started)
