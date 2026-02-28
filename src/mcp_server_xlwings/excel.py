"""xlwings COM automation wrapper for Excel manipulation."""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any

import xlwings as xw


class ExcelError(Exception):
    """Human-readable Excel operation error."""


def _serialize_value(value: Any) -> Any:
    """Convert Excel cell values to JSON-serializable types."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, datetime.time):
        return value.isoformat()
    if isinstance(value, float):
        if value == int(value):
            return int(value)
        return value
    return value


def _serialize_row(row: list[Any]) -> list[Any]:
    return [_serialize_value(v) for v in row]


def _to_2d(raw: Any) -> list[list[Any]]:
    """Normalise raw xlwings value to a serialized 2D list."""
    if raw is None:
        return []
    if not isinstance(raw, list):
        return [[_serialize_value(raw)]]
    if raw and not isinstance(raw[0], list):
        return [_serialize_row(raw)]
    return [_serialize_row(row) for row in raw]


class ExcelHandler:
    """Stateless wrapper around xlwings for Excel COM automation."""

    # ------------------------------------------------------------------ #
    #  Internal helpers
    # ------------------------------------------------------------------ #

    def _get_app(self) -> xw.App:
        """Return the running Excel App instance, or create one."""
        try:
            app = xw.apps.active
            if app is None:
                app = xw.App(visible=True)
            return app
        except Exception:
            return xw.App(visible=True)

    def _get_workbook(self, workbook: str) -> xw.Book:
        """Find an open workbook by name or full path, or open it."""
        app = self._get_app()
        for wb in app.books:
            if wb.name == workbook or wb.fullname == workbook:
                return wb
        path = Path(workbook)
        if path.exists():
            return app.books.open(str(path))
        raise ExcelError(
            f"Workbook '{workbook}' is not open and the path does not exist."
        )

    def _get_workbook_or_active(self, workbook: str | None) -> xw.Book:
        """Return the specified workbook, or the active workbook if None."""
        if workbook is not None:
            return self._get_workbook(workbook)
        app = self._get_app()
        wb = app.books.active
        if wb is None:
            raise ExcelError("No active workbook found.")
        return wb

    def _get_sheet(self, wb: xw.Book, sheet: str | None) -> xw.Sheet:
        """Return the requested sheet or the active sheet."""
        if sheet is None:
            return wb.sheets.active
        try:
            return wb.sheets[sheet]
        except Exception:
            names = [s.name for s in wb.sheets]
            raise ExcelError(
                f"Sheet '{sheet}' not found. Available sheets: {names}"
            )

    # ------------------------------------------------------------------ #
    #  Tool 1: get_active_workbook (includes selection data)
    # ------------------------------------------------------------------ #

    def get_active_workbook(self) -> dict:
        app = self._get_app()
        wb = app.books.active
        if wb is None:
            raise ExcelError("No active workbook found.")

        ws = wb.sheets.active
        sheets_meta = []
        for s in wb.sheets:
            used = s.used_range
            sheets_meta.append({
                "name": s.name,
                "used_range": used.address,
                "rows": used.rows.count,
                "columns": used.columns.count,
            })
        info: dict[str, Any] = {
            "name": wb.name,
            "path": wb.fullname,
            "sheets": sheets_meta,
            "active_sheet": ws.name,
        }

        try:
            sel = app.selection
            if sel is not None:
                sel_data = _to_2d(sel.value)
                info["selection"] = {
                    "address": sel.address,
                    "sheet": sel.sheet.name,
                    "data": sel_data,
                    "rows": len(sel_data),
                    "columns": len(sel_data[0]) if sel_data else 0,
                }
        except Exception:
            pass

        return info

    # ------------------------------------------------------------------ #
    #  Tool 2: manage_workbooks
    # ------------------------------------------------------------------ #

    def manage_workbooks(
        self,
        action: str,
        workbook: str | None = None,
        filepath: str | None = None,
        read_only: bool = True,
        save: bool = False,
    ) -> dict | list[dict]:
        if action == "list":
            app = self._get_app()
            result = []
            for wb in app.books:
                sheets_meta = []
                for s in wb.sheets:
                    used = s.used_range
                    sheets_meta.append({
                        "name": s.name,
                        "used_range": used.address,
                        "rows": used.rows.count,
                        "columns": used.columns.count,
                    })
                result.append({
                    "name": wb.name,
                    "path": wb.fullname,
                    "sheets": sheets_meta,
                })
            return result

        if action == "open":
            if filepath is None:
                raise ExcelError(
                    "Parameter 'filepath' is required for open action. "
                    "Use 'new' to create a blank workbook."
                )
            app = self._get_app()
            if filepath.lower() == "new":
                wb = app.books.add()
            else:
                path = Path(filepath)
                if not path.exists():
                    raise ExcelError(f"File not found: {filepath}")
                for wb in app.books:
                    if wb.fullname == str(path.resolve()):
                        return {
                            "name": wb.name,
                            "path": wb.fullname,
                            "sheets": [s.name for s in wb.sheets],
                        }
                wb = app.books.open(str(path), read_only=read_only)
            return {
                "name": wb.name,
                "path": wb.fullname,
                "sheets": [s.name for s in wb.sheets],
            }

        if action == "save":
            wb = self._get_workbook_or_active(workbook)
            if filepath:
                wb.save(filepath)
                return {"message": f"Workbook saved as '{filepath}'."}
            wb.save()
            return {"message": f"Workbook '{wb.name}' saved to '{wb.fullname}'."}

        if action == "close":
            wb = self._get_workbook_or_active(workbook)
            name = wb.name
            if save:
                wb.save()
            wb.close()
            return {"message": f"Workbook '{name}' closed (save={save})."}

        if action == "recalculate":
            app = self._get_app()
            if workbook:
                wb = self._get_workbook(workbook)
                wb.app.calculate()
                return {"message": f"Workbook '{wb.name}' recalculated."}
            app.calculate()
            return {"message": "All open workbooks recalculated."}

        raise ExcelError(
            f"Unknown action '{action}'. Use: list, open, save, close, recalculate."
        )

    # ------------------------------------------------------------------ #
    #  Tool 3: read_data
    # ------------------------------------------------------------------ #

    def read_data(
        self,
        workbook: str | None = None,
        sheet: str | None = None,
        cell_range: str | None = None,
        headers: bool = True,
        detail: bool = False,
    ) -> dict:
        wb = self._get_workbook_or_active(workbook)
        ws = self._get_sheet(wb, sheet)

        # No range specified: return sheet summary without reading data
        if cell_range is None:
            used = ws.used_range
            total_rows = used.rows.count
            total_cols = used.columns.count
            result: dict[str, Any] = {
                "sheet": ws.name,
                "used_range": used.address,
                "total_rows": total_rows,
                "total_columns": total_cols,
            }
            # Include first-row headers for context
            if total_rows >= 1:
                first_row = ws.range(
                    (used.row, used.column),
                    (used.row, used.column + total_cols - 1),
                ).value
                if not isinstance(first_row, list):
                    first_row = [first_row]
                result["headers"] = [_serialize_value(v) for v in first_row]

            # Merged cells - scan header area (first 20 rows) for performance
            try:
                merge_flag = ws.api.UsedRange.MergeCells
                # False = no merges, True/None = some or all merged
                if merge_flag is not False:
                    merged: list[dict] = []
                    visited_addrs: set[str] = set()
                    scan_rows = min(20, total_rows)
                    header_rng = ws.range(
                        (used.row, used.column),
                        (used.row + scan_rows - 1, used.column + total_cols - 1),
                    )
                    for cell in header_rng:
                        if cell.api.MergeCells:
                            ma = cell.api.MergeArea
                            addr = ma.Address
                            if addr not in visited_addrs:
                                visited_addrs.add(addr)
                                merged.append({
                                    "range": addr,
                                    "value": _serialize_value(ma.Cells(1, 1).Value),
                                    "rows": ma.Rows.Count,
                                    "columns": ma.Columns.Count,
                                })
                    if merged:
                        result["merged_cells"] = merged
            except Exception:
                pass

            # Region detection - probe strategic cells (fast, max ~5 COM calls)
            try:
                api_used = ws.api.UsedRange
                first_cr = api_used.Cells(1, 1).CurrentRegion
                regions: list[dict] = [{
                    "range": first_cr.Address,
                    "rows": first_cr.Rows.Count,
                    "columns": first_cr.Columns.Count,
                }]
                seen: set[str] = {first_cr.Address}

                for probe_row in [total_rows, total_rows // 2]:
                    if probe_row < 2:
                        continue
                    probe = api_used.Cells(probe_row, 1)
                    if probe.Value is not None:
                        cr = probe.CurrentRegion
                        if cr.Address not in seen:
                            seen.add(cr.Address)
                            regions.append({
                                "range": cr.Address,
                                "rows": cr.Rows.Count,
                                "columns": cr.Columns.Count,
                            })
                if len(regions) > 1:
                    result["regions"] = regions
            except Exception:
                pass

            return result

        rng = ws.range(cell_range)
        raw = rng.value
        if raw is None:
            return {
                "data": [],
                "range": rng.address,
                "sheet": ws.name,
                "rows": 0,
                "columns": 0,
            }

        data = _to_2d(raw)

        result: dict[str, Any] = {
            "range": rng.address,
            "sheet": ws.name,
            "rows": len(data),
            "columns": len(data[0]) if data else 0,
        }

        if headers and len(data) > 1:
            result["headers"] = data[0]
            result["data"] = data[1:]
        else:
            result["data"] = data

        if detail and rng.size == 1:
            raw_val = rng.value
            formula = rng.formula if rng.formula != rng.value else None

            if raw_val is None:
                value_type = "empty"
            elif isinstance(raw_val, bool):
                value_type = "boolean"
            elif isinstance(raw_val, (int, float)):
                value_type = "number"
            elif isinstance(raw_val, str):
                value_type = "text"
            elif isinstance(raw_val, (datetime.datetime, datetime.date)):
                value_type = "date"
            else:
                value_type = type(raw_val).__name__

            result["detail"] = {
                "value": _serialize_value(raw_val),
                "type": value_type,
                "formula": formula,
                "number_format": rng.number_format,
            }
            try:
                font = rng.font
                result["detail"]["font"] = {
                    "name": font.name,
                    "size": font.size,
                    "bold": font.bold,
                    "italic": font.italic,
                    "color": font.color,
                }
            except Exception:
                pass

        return result

    # ------------------------------------------------------------------ #
    #  Tool 4: write_data
    # ------------------------------------------------------------------ #

    def write_data(
        self,
        start_cell: str,
        data: list[list] | None = None,
        formula: str | None = None,
        workbook: str | None = None,
        sheet: str | None = None,
    ) -> dict:
        if data is not None and formula is not None:
            raise ExcelError("Provide either 'data' or 'formula', not both.")
        if data is None and formula is None:
            raise ExcelError("Either 'data' or 'formula' must be provided.")

        wb = self._get_workbook_or_active(workbook)
        ws = self._get_sheet(wb, sheet)

        if formula is not None:
            rng = ws.range(start_cell)
            rng.formula = formula
            calculated = _serialize_value(rng.value)
            return {
                "cell": rng.address,
                "sheet": ws.name,
                "formula": formula,
                "calculated_value": calculated,
            }

        rng = ws.range(start_cell)
        rng.value = data
        rows = len(data)
        cols = len(data[0]) if data else 0
        end_cell = rng.expand().address if rows and cols else rng.address
        return {
            "message": f"Data written successfully to {ws.name}",
            "start_cell": start_cell,
            "written_range": end_cell,
            "rows": rows,
            "columns": cols,
        }

    # ------------------------------------------------------------------ #
    #  Tool 5: manage_sheets
    # ------------------------------------------------------------------ #

    def manage_sheets(
        self,
        action: str,
        workbook: str | None = None,
        sheet: str | None = None,
        new_name: str | None = None,
        position: int = 1,
        count: int = 1,
    ) -> dict:
        wb = self._get_workbook_or_active(workbook)

        if action == "list":
            return {
                "sheets": [s.name for s in wb.sheets],
                "active_sheet": wb.sheets.active.name,
            }

        if action == "add":
            name = new_name or sheet
            ws = wb.sheets.add(after=wb.sheets[wb.sheets.count - 1])
            if name:
                ws.name = name
            return {
                "message": f"Sheet '{ws.name}' added.",
                "sheets": [s.name for s in wb.sheets],
            }

        if action == "delete":
            if sheet is None:
                raise ExcelError("Parameter 'sheet' is required for delete action.")
            app = wb.app
            app.display_alerts = False
            try:
                wb.sheets[sheet].delete()
            finally:
                app.display_alerts = True
            return {
                "message": f"Sheet '{sheet}' deleted.",
                "sheets": [s.name for s in wb.sheets],
            }

        if action == "rename":
            if sheet is None or new_name is None:
                raise ExcelError(
                    "Both 'sheet' and 'new_name' are required for rename action."
                )
            wb.sheets[sheet].name = new_name
            return {
                "message": f"Sheet '{sheet}' renamed to '{new_name}'.",
                "sheets": [s.name for s in wb.sheets],
            }

        if action == "copy":
            if sheet is None:
                raise ExcelError("Parameter 'sheet' is required for copy action.")
            source = wb.sheets[sheet]
            source.copy(after=source)
            copied = wb.sheets[source.index + 1]
            if new_name:
                copied.name = new_name
            return {
                "message": f"Sheet '{sheet}' copied as '{copied.name}'.",
                "sheets": [s.name for s in wb.sheets],
            }

        if action == "activate":
            if sheet is None:
                raise ExcelError("Parameter 'sheet' is required for activate action.")
            ws = self._get_sheet(wb, sheet)
            ws.activate()
            return {
                "message": f"Sheet '{sheet}' is now active.",
                "workbook": wb.name,
                "active_sheet": ws.name,
            }

        # Row/column insert/delete actions
        if action in ("insert_rows", "delete_rows", "insert_columns", "delete_columns"):
            ws = self._get_sheet(wb, sheet)
            if count < 1:
                raise ExcelError("count must be >= 1.")

            if action == "insert_rows":
                rng_str = f"{position}:{position + count - 1}"
                ws.range(rng_str).api.EntireRow.Insert()
                return {"message": f"Inserted {count} row(s) at row {position}.", "sheet": ws.name}

            if action == "delete_rows":
                rng_str = f"{position}:{position + count - 1}"
                ws.range(rng_str).api.EntireRow.Delete()
                return {"message": f"Deleted {count} row(s) starting at row {position}.", "sheet": ws.name}

            if action == "insert_columns":
                rng = ws.range((1, position), (1, position + count - 1))
                rng.api.EntireColumn.Insert()
                return {"message": f"Inserted {count} column(s) at column {position}.", "sheet": ws.name}

            if action == "delete_columns":
                rng = ws.range((1, position), (1, position + count - 1))
                rng.api.EntireColumn.Delete()
                return {"message": f"Deleted {count} column(s) starting at column {position}.", "sheet": ws.name}

        raise ExcelError(
            f"Unknown action '{action}'. Use: list, add, delete, rename, copy, "
            "activate, insert_rows, delete_rows, insert_columns, delete_columns."
        )

    # ------------------------------------------------------------------ #
    #  Tool 6: find_replace
    # ------------------------------------------------------------------ #

    def find_replace(
        self,
        find: str,
        workbook: str | None = None,
        sheet: str | None = None,
        replace: str | None = None,
        match_case: bool = False,
    ) -> dict:
        wb = self._get_workbook_or_active(workbook)
        ws = self._get_sheet(wb, sheet)

        rng = ws.used_range
        if rng is None:
            return {"matches": [], "count": 0}

        api = rng.api

        if replace is not None:
            count = api.Replace(
                What=find,
                Replacement=replace,
                MatchCase=match_case,
            )
            return {
                "message": f"Replaced occurrences of '{find}' with '{replace}'.",
                "replaced": bool(count),
            }

        matches: list[dict] = []
        first = api.Find(What=find, MatchCase=match_case)
        if first is None:
            return {"matches": [], "count": 0}

        current = first
        while True:
            matches.append({
                "cell": current.Address,
                "value": _serialize_value(current.Value),
                "sheet": ws.name,
            })
            current = api.FindNext(current)
            if current is None or current.Address == first.Address:
                break

        return {"matches": matches, "count": len(matches)}

    # ------------------------------------------------------------------ #
    #  Tool 7: format_range
    # ------------------------------------------------------------------ #

    def format_range(
        self,
        cell_range: str,
        workbook: str | None = None,
        sheet: str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        font_size: int | None = None,
        font_color: str | None = None,
        bg_color: str | None = None,
        number_format: str | None = None,
        alignment: str | None = None,
        wrap_text: bool | None = None,
        border: bool | None = None,
    ) -> dict:
        wb = self._get_workbook_or_active(workbook)
        ws = self._get_sheet(wb, sheet)
        rng = ws.range(cell_range)

        applied: list[str] = []

        if bold is not None:
            rng.font.bold = bold
            applied.append(f"bold={bold}")

        if italic is not None:
            rng.font.italic = italic
            applied.append(f"italic={italic}")

        if underline is not None:
            rng.api.Font.Underline = 2 if underline else -4142
            applied.append(f"underline={underline}")

        if font_size is not None:
            rng.font.size = font_size
            applied.append(f"font_size={font_size}")

        if font_color is not None:
            rng.font.color = font_color
            applied.append(f"font_color={font_color}")

        if bg_color is not None:
            rng.color = bg_color
            applied.append(f"bg_color={bg_color}")

        if number_format is not None:
            rng.number_format = number_format
            applied.append(f"number_format={number_format}")

        if alignment is not None:
            align_map = {
                "left": -4131,
                "center": -4108,
                "right": -4152,
                "justify": -4130,
            }
            xl_align = align_map.get(alignment.lower())
            if xl_align is None:
                raise ExcelError(
                    f"Unknown alignment '{alignment}'. "
                    "Use: left, center, right, justify."
                )
            rng.api.HorizontalAlignment = xl_align
            applied.append(f"alignment={alignment}")

        if wrap_text is not None:
            rng.wrap_text = wrap_text
            applied.append(f"wrap_text={wrap_text}")

        if border is not None and border:
            for edge in range(7, 13):
                try:
                    rng.api.Borders(edge).LineStyle = 1
                    rng.api.Borders(edge).Weight = 2
                except Exception:
                    pass
            applied.append("border=True")

        return {
            "range": rng.address,
            "sheet": ws.name,
            "applied": applied,
        }

    # ------------------------------------------------------------------ #
    #  Tool 8: run_macro
    # ------------------------------------------------------------------ #

    def run_macro(
        self,
        macro_name: str,
        workbook: str | None = None,
        args: list | None = None,
    ) -> dict:
        app = self._get_app()

        if workbook:
            wb = self._get_workbook(workbook)
            qualified = f"'{wb.name}'!{macro_name}"
        else:
            qualified = macro_name

        try:
            if args:
                result = app.macro(qualified)(*args)
            else:
                result = app.macro(qualified)()
        except Exception as exc:
            raise ExcelError(f"Failed to run macro '{macro_name}': {exc}")

        return {
            "message": f"Macro '{macro_name}' executed successfully.",
            "return_value": _serialize_value(result),
        }
