"""MCP transport layer for xl-agent-core."""


def create_mcp_server(*args, **kwargs):
    from xl_agent_core.mcp.server import create_mcp_server as _create_mcp_server

    return _create_mcp_server(*args, **kwargs)


__all__ = ["create_mcp_server"]
