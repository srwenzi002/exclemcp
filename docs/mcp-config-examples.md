# MCP Config Examples (Docker)

Image name used below: `excel-mcp:latest`

Before client config:

```bash
docker build -t excel-mcp:latest .
```

## Codex CLI

Recommended: auto-bind current working directory (`$PWD`) so you can use this MCP in any project directory.

```bash
codex mcp add excel-local \
  -- zsh -lc 'docker run -i --rm \
  -v "$PWD:$PWD" \
  -w "$PWD" \
  -e EXCEL_MCP_ROOT="$PWD" \
  excel-mcp:latest'
```

Optional: fixed shared directory mode.

```bash
mkdir -p ~/excel-mcp-data
codex mcp add excel-local-fixed \
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
      "command": "zsh",
      "args": [
        "-lc",
        "docker run -i --rm -v \"$PWD:$PWD\" -w \"$PWD\" -e EXCEL_MCP_ROOT=\"$PWD\" excel-mcp:latest"
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
      "command": "zsh",
      "args": [
        "-lc",
        "docker run -i --rm -v \"$PWD:$PWD\" -w \"$PWD\" -e EXCEL_MCP_ROOT=\"$PWD\" excel-mcp:latest"
      ]
    }
  }
}
```
