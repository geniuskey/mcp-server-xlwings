import { defineConfig } from "vitepress";

const pypiIcon =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M12 2 L22 7.5 L12 13 L2 7.5Z" fill="#4B8BBE"/><path d="M2 7.5 L12 13 L12 23 L2 17.5Z" fill="#3775A9"/><path d="M22 7.5 L12 13 L12 23 L22 17.5Z" fill="#006DAD"/></svg>';

export default defineConfig({
  title: "mcp-server-xlwings",
  description:
    "MCP server for Excel automation via xlwings COM. Works with DRM-protected files.",
  base: "/mcp-server-xlwings/",
  head: [
    [
      "link",
      {
        rel: "icon",
        type: "image/svg+xml",
        href: "/mcp-server-xlwings/favicon.svg",
      },
    ],
    [
      "meta",
      { property: "og:title", content: "mcp-server-xlwings" },
    ],
    [
      "meta",
      {
        property: "og:description",
        content:
          "MCP server for Excel automation via xlwings COM. Works with DRM-protected files.",
      },
    ],
  ],
  locales: {
    root: {
      label: "English",
      lang: "en",
    },
    ko: {
      label: "한국어",
      lang: "ko",
      description:
        "xlwings COM을 통한 Excel 자동화 MCP 서버. DRM 보호 파일 지원.",
      themeConfig: {
        nav: [
          { text: "가이드", link: "/ko/guide/getting-started" },
          { text: "도구", link: "/ko/tools/" },
          { text: "예시", link: "/ko/examples" },
        ],
        sidebar: [
          {
            text: "가이드",
            items: [
              { text: "시작하기", link: "/ko/guide/getting-started" },
              { text: "설정", link: "/ko/guide/configuration" },
              { text: "아키텍처", link: "/ko/guide/architecture" },
              { text: "비교", link: "/ko/guide/comparison" },
              { text: "성능 가이드", link: "/ko/guide/performance" },
              { text: "문제 해결", link: "/ko/guide/troubleshooting" },
            ],
          },
          {
            text: "레퍼런스",
            items: [
              { text: "도구", link: "/ko/tools/" },
              { text: "예시", link: "/ko/examples" },
              { text: "변경 이력", link: "/ko/changelog" },
            ],
          },
        ],
      },
    },
  },
  themeConfig: {
    logo: "/logo.svg",
    nav: [
      { text: "Guide", link: "/guide/getting-started" },
      { text: "Tools", link: "/tools/" },
      { text: "Examples", link: "/examples" },
    ],
    sidebar: [
      {
        text: "Guide",
        items: [
          { text: "Getting Started", link: "/guide/getting-started" },
          { text: "Configuration", link: "/guide/configuration" },
          { text: "Architecture", link: "/guide/architecture" },
          { text: "Comparison", link: "/guide/comparison" },
          { text: "Performance", link: "/guide/performance" },
          { text: "Troubleshooting", link: "/guide/troubleshooting" },
        ],
      },
      {
        text: "Reference",
        items: [
          { text: "Tools", link: "/tools/" },
          { text: "Examples", link: "/examples" },
          { text: "Changelog", link: "/changelog" },
        ],
      },
    ],
    socialLinks: [
      {
        icon: "github",
        link: "https://github.com/geniuskey/mcp-server-xlwings",
      },
      {
        icon: { svg: pypiIcon },
        link: "https://pypi.org/project/mcp-server-xlwings/",
        ariaLabel: "PyPI",
      },
    ],
    footer: {
      message: "Released under the MIT License.",
    },
    search: {
      provider: "local",
    },
  },
});
