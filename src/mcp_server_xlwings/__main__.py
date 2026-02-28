"""Entry point for ``python -m mcp_server_xlwings``."""

from .server import mcp


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
