import { defineConfig } from "vitepress";

export default defineConfig({
  title: "mcp-server-xlwings",
  description:
    "MCP server for Excel automation via xlwings COM. Works with DRM-protected files.",
  base: "/mcp-server-xlwings/",
  head: [
    ["link", { rel: "icon", type: "image/svg+xml", href: "/mcp-server-xlwings/favicon.svg" }],
    [
      "meta",
      {
        property: "og:title",
        content: "mcp-server-xlwings",
      },
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
        ],
      },
      {
        text: "Reference",
        items: [
          { text: "Tools", link: "/tools/" },
          { text: "Examples", link: "/examples" },
        ],
      },
    ],
    socialLinks: [
      {
        icon: "github",
        link: "https://github.com/geniuskey/mcp-server-xlwings",
      },
      {
        icon: { svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><path d="M431.5 80.5h-351c-22.1 0-40 17.9-40 40v271c0 22.1 17.9 40 40 40h351c22.1 0 40-17.9 40-40v-271c0-22.1-17.9-40-40-40zm-188 305h-65l-16.4-67h-28.6v67h-53v-231h97.6c43.4 0 70.4 24.4 70.4 65.6 0 30.2-14.6 51.6-39 60.8l38 104.6zm109 0h-53v-231h53v231zm98 0h-53v-231h53v231z"/><path d="M228.1 199.5c0-14.6-9-22.6-24.6-22.6h-40v46.6h40c15.6 0 24.6-8.6 24.6-24z"/></svg>' },
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
