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
        icon: { svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M12 2 L22 7.5 L12 13 L2 7.5Z" fill="#4B8BBE"/><path d="M2 7.5 L12 13 L12 23 L2 17.5Z" fill="#3775A9"/><path d="M22 7.5 L12 13 L12 23 L22 17.5Z" fill="#006DAD"/></svg>' },
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
