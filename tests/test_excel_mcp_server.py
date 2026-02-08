from __future__ import annotations

from pathlib import Path

import pytest

from excel_mcp_server import (
    clear_range,
    delete_columns,
    delete_rows,
    delete_sheet,
    format_range,
    insert_columns,
    insert_rows,
    list_sheets,
    read_range,
    rename_sheet,
    write_cell,
    write_range,
)


def _p(path: Path) -> str:
    return str(path.resolve())


def test_security_workspace_boundary(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("EXCEL_MCP_ROOT", _p(tmp_path))
    outside = tmp_path.parent / "outside.xlsx"
    with pytest.raises(ValueError):
        list_sheets(_p(outside))


def test_security_extension(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("EXCEL_MCP_ROOT", _p(tmp_path))
    bad = tmp_path / "bad.xls"
    with pytest.raises(ValueError):
        list_sheets(_p(bad), create_if_missing=True)


def test_read_write_and_structural_ops(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("EXCEL_MCP_ROOT", _p(tmp_path))
    file_path = _p(tmp_path / "demo.xlsx")

    write_range(
        file_path=file_path,
        sheet_name="Data",
        start_cell="A1",
        values=[["name", "value"], ["gpu", 123], ["cpu", 456]],
        create_if_missing=True,
    )
    insert_rows(file_path=file_path, sheet_name="Data", idx=2, amount=1)
    write_cell(file_path=file_path, sheet_name="Data", cell="A2", value="inserted")
    insert_columns(file_path=file_path, sheet_name="Data", idx=2, amount=1)
    write_cell(file_path=file_path, sheet_name="Data", cell="B1", value="new_col")
    delete_rows(file_path=file_path, sheet_name="Data", idx=4, amount=1)
    delete_columns(file_path=file_path, sheet_name="Data", idx=3, amount=1)

    values = read_range(file_path=file_path, sheet_name="Data", cell_range="A1:C4")["values"]
    assert values[0][0] == "name"
    assert values[0][1] == "new_col"
    assert values[1][0] == "inserted"


def test_sheet_and_format_ops(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("EXCEL_MCP_ROOT", _p(tmp_path))
    file_path = _p(tmp_path / "format.xlsx")

    write_range(
        file_path=file_path,
        sheet_name="SheetA",
        start_cell="A1",
        values=[["header"], ["body"]],
        create_if_missing=True,
    )
    rename_sheet(file_path=file_path, old_name="SheetA", new_name="Report")
    write_cell(file_path=file_path, sheet_name="Extra", cell="A1", value="x", create_if_missing=True)
    clear_range(file_path=file_path, sheet_name="Report", cell_range="A2:A2")
    format_range(
        file_path=file_path,
        sheet_name="Report",
        cell_range="A1:A1",
        bold=True,
        wrap_text=True,
        horizontal="center",
        fill_hex="EAF2FF",
    )
    delete_sheet(file_path=file_path, sheet_name="Extra")

    sheets = list_sheets(file_path=file_path)["sheets"]
    assert "Report" in sheets
    assert "Extra" not in sheets
    assert read_range(file_path=file_path, sheet_name="Report", cell_range="A2:A2")["values"][0][0] is None
