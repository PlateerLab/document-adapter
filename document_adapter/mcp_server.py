"""MCP stdio server — Claude Desktop/Code에서 document-adapter tool을 호출하게 해준다.

실행:
    python -m document_adapter.mcp_server

Claude Desktop 설정 예시 (~/Library/Application Support/Claude/claude_desktop_config.json):
{
  "mcpServers": {
    "document-adapter": {
      "command": "/path/to/venv/bin/python",
      "args": ["-m", "document_adapter.mcp_server"]
    }
  }
}
"""
from __future__ import annotations

import asyncio
import json
import logging

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from .tools import TOOL_DEFINITIONS, call_tool

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("document-adapter-mcp")

server: Server = Server("document-adapter")


@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name=t["name"],
            description=t["description"],
            inputSchema=t["input_schema"],
        )
        for t in TOOL_DEFINITIONS
    ]


@server.call_tool()
async def on_call_tool(name: str, arguments: dict) -> list[TextContent]:
    log.info("tool call: %s %s", name, list(arguments.keys()))
    result = call_tool(name, arguments)
    return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]


async def main() -> None:
    async with stdio_server() as (read, write):
        await server.run(read, write, server.create_initialization_options())


def main_sync() -> None:
    """Console script entry point."""
    asyncio.run(main())


if __name__ == "__main__":
    main_sync()
