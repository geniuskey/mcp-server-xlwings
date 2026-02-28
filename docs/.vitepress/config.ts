import { defineConfig } from "vitepress";

export default defineConfig({
  title: "mcp-server-xlwings",
  description:
    "MCP server for Excel automation via xlwings COM. Works with DRM-protected files.",
  base: "/mcp-server-xlwings/",
  head: [
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
    ],
    footer: {
      message: "Released under the MIT License.",
    },
    search: {
      provider: "local",
    },
  },
});
