# MCP Config Examples (Docker)

Image name used below: `excel-mcp:latest`

Before client config:

```bash
docker build -t excel-mcp:latest .
mkdir -p ~/excel-mcp-data
```

## Codex CLI

```bash
codex mcp add excel-local \
  --env EXCEL_MCP_ROOT=/workspace \
  -- docker run -i --rm \
  -v "$HOME/excel-mcp-data:/workspace" \
  excel-mcp:latest
```

Verify:

```bash
codex mcp list
codex mcp get excel-local
```

## Cursor

Add to Cursor MCP config:

```json
{
  "mcpServers": {
    "excel-local": {
      "command": "docker",
      "args": [
        "run",
        "-i",
        "--rm",
        "-v",
        "/Users/<you>/excel-mcp-data:/workspace",
        "-e",
        "EXCEL_MCP_ROOT=/workspace",
        "excel-mcp:latest"
      ]
    }
  }
}
```

## GitHub Copilot in VS Code

Create `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "excel-local": {
      "type": "stdio",
      "command": "docker",
      "args": [
        "run",
        "-i",
        "--rm",
        "-v",
        "/Users/<you>/excel-mcp-data:/workspace",
        "-e",
        "EXCEL_MCP_ROOT=/workspace",
        "excel-mcp:latest"
      ]
    }
  }
}
```
