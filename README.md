# excel-mcp

支持读取和编辑 Excel（`.xlsx/.xlsm`）的 MCP Server。

## 功能

- `list_sheets`
- `read_range`
- `write_cell`
- `write_range`
- `insert_rows` / `delete_rows`
- `insert_columns` / `delete_columns`
- `rename_sheet` / `delete_sheet`
- `clear_range`
- `format_range`

## 安全策略

- 默认仅允许访问当前工作目录下文件。
- 可通过 `EXCEL_MCP_ROOT` 指定允许访问的根目录。
- 仅允许 `.xlsx` 和 `.xlsm`。

## 本地开发运行

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
pytest -q
excel-mcp
```

## Docker 运行

```bash
docker build -t excel-mcp:latest .
mkdir -p ~/excel-mcp-data
docker run -i --rm \
  -v "$HOME/excel-mcp-data:/workspace" \
  -e EXCEL_MCP_ROOT=/workspace \
  excel-mcp:latest
```

## 客户端配置（Codex/Cursor/Copilot）

见 `docs/mcp-config-examples.md`。

## 发布到 Git（首次）

```bash
git init
git add .
git commit -m "feat: dockerized excel mcp with client configs"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```
