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
        # Not found among open books -- try to open from path
        path = Path(workbook)
        if path.exists():
            return app.books.open(str(path))
        raise ExcelError(
            f"Workbook '{workbook}' is not open and the path does not exist."
        )

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
    #  Tool implementations
    # ------------------------------------------------------------------ #

    def list_open_workbooks(self) -> list[dict]:
        app = self._get_app()
        result = []
        for wb in app.books:
            result.append(
                {
                    "name": wb.name,
                    "path": wb.fullname,
                    "sheets": [s.name for s in wb.sheets],
                }
            )
        return result

    def open_workbook(
        self, filepath: str, read_only: bool = True
    ) -> dict:
        app = self._get_app()
        if filepath.lower() == "new":
            wb = app.books.add()
        else:
            path = Path(filepath)
            if not path.exists():
                raise ExcelError(f"File not found: {filepath}")
            # Check if already open
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

    def read_range(
        self,
        workbook: str,
        sheet: str | None = None,
        cell_range: str | None = None,
        headers: bool = True,
    ) -> dict:
        wb = self._get_workbook(workbook)
        ws = self._get_sheet(wb, sheet)

        if cell_range:
            rng = ws.range(cell_range)
        else:
            rng = ws.used_range

        raw = rng.value
        if raw is None:
            return {
                "data": [],
                "range": rng.address,
                "sheet": ws.name,
                "rows": 0,
                "columns": 0,
            }

        # Normalise to 2D list
        if not isinstance(raw, list):
            data = [[raw]]
        elif raw and not isinstance(raw[0], list):
            data = [raw]
        else:
            data = raw

        data = [_serialize_row(row) for row in data]

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

        return result

    def write_range(
        self,
        workbook: str,
        sheet: str | None,
        start_cell: str,
        data: list[list],
    ) -> dict:
        wb = self._get_workbook(workbook)
        ws = self._get_sheet(wb, sheet)

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

    def read_cell_info(
        self,
        workbook: str,
        sheet: str | None,
        cell: str,
    ) -> dict:
        wb = self._get_workbook(workbook)
        ws = self._get_sheet(wb, sheet)
        rng = ws.range(cell)

        value = _serialize_value(rng.value)
        formula = rng.formula if rng.formula != rng.value else None

        # Determine type label
        raw = rng.value
        if raw is None:
            value_type = "empty"
        elif isinstance(raw, bool):
            value_type = "boolean"
        elif isinstance(raw, (int, float)):
            value_type = "number"
        elif isinstance(raw, str):
            value_type = "text"
        elif isinstance(raw, (datetime.datetime, datetime.date)):
            value_type = "date"
        else:
            value_type = type(raw).__name__

        info: dict[str, Any] = {
            "cell": rng.address,
            "sheet": ws.name,
            "value": value,
            "type": value_type,
            "formula": formula,
            "number_format": rng.number_format,
        }

        # Font info
        try:
            font = rng.font
            info["font"] = {
                "name": font.name,
                "size": font.size,
                "bold": font.bold,
                "italic": font.italic,
                "color": font.color,
            }
        except Exception:
            pass

        return info

    def set_formula(
        self,
        workbook: str,
        sheet: str | None,
        cell: str,
        formula: str,
    ) -> dict:
        wb = self._get_workbook(workbook)
        ws = self._get_sheet(wb, sheet)
        rng = ws.range(cell)

        rng.formula = formula
        # Read back the calculated value
        calculated = _serialize_value(rng.value)

        return {
            "cell": rng.address,
            "sheet": ws.name,
            "formula": formula,
            "calculated_value": calculated,
        }

    def manage_sheets(
        self,
        workbook: str,
        action: str,
        sheet: str | None = None,
        new_name: str | None = None,
    ) -> dict:
        wb = self._get_workbook(workbook)

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
            return {"message": f"Sheet '{ws.name}' added.", "sheets": [s.name for s in wb.sheets]}

        if action == "delete":
            if sheet is None:
                raise ExcelError("Parameter 'sheet' is required for delete action.")
            app = wb.app
            app.display_alerts = False
            try:
                wb.sheets[sheet].delete()
            finally:
                app.display_alerts = True
            return {"message": f"Sheet '{sheet}' deleted.", "sheets": [s.name for s in wb.sheets]}

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

        raise ExcelError(
            f"Unknown action '{action}'. Use: list, add, delete, rename, copy."
        )

    def find_replace(
        self,
        workbook: str,
        sheet: str | None,
        find: str,
        replace: str | None = None,
        match_case: bool = False,
    ) -> dict:
        wb = self._get_workbook(workbook)
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

        # Search only -- collect matching cells
        matches: list[dict] = []
        first = api.Find(What=find, MatchCase=match_case)
        if first is None:
            return {"matches": [], "count": 0}

        current = first
        while True:
            matches.append(
                {
                    "cell": current.Address,
                    "value": _serialize_value(current.Value),
                    "sheet": ws.name,
                }
            )
            current = api.FindNext(current)
            if current is None or current.Address == first.Address:
                break

        return {"matches": matches, "count": len(matches)}

    def format_range(
        self,
        workbook: str,
        sheet: str | None,
        cell_range: str,
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
        wb = self._get_workbook(workbook)
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
            # xlUnderlineStyleSingle = 2, xlUnderlineStyleNone = -4142
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
                "left": -4131,      # xlLeft
                "center": -4108,    # xlCenter
                "right": -4152,     # xlRight
                "justify": -4130,   # xlJustify
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
            # Apply thin borders on all edges (xlThin = 2)
            for edge in range(7, 13):  # xlEdgeLeft(7) through xlInsideHorizontal(12)
                try:
                    rng.api.Borders(edge).LineStyle = 1  # xlContinuous
                    rng.api.Borders(edge).Weight = 2  # xlThin
                except Exception:
                    pass
            applied.append("border=True")

        return {
            "range": rng.address,
            "sheet": ws.name,
            "applied": applied,
        }

    def save_workbook(
        self,
        workbook: str,
        filepath: str | None = None,
    ) -> dict:
        wb = self._get_workbook(workbook)

        if filepath:
            wb.save(filepath)
            return {"message": f"Workbook saved as '{filepath}'."}

        wb.save()
        return {"message": f"Workbook '{wb.name}' saved to '{wb.fullname}'."}

    # ------------------------------------------------------------------ #
    #  Active Excel tools (xlwings-exclusive capabilities)
    # ------------------------------------------------------------------ #

    def get_active_workbook(self) -> dict:
        app = self._get_app()
        wb = app.books.active
        if wb is None:
            raise ExcelError("No active workbook found.")

        ws = wb.sheets.active
        info: dict[str, Any] = {
            "name": wb.name,
            "path": wb.fullname,
            "sheets": [s.name for s in wb.sheets],
            "active_sheet": ws.name,
        }

        # Include current selection info
        try:
            sel = app.selection
            if sel is not None:
                info["selection"] = {
                    "address": sel.address,
                    "sheet": sel.sheet.name,
                }
        except Exception:
            pass

        return info

    def read_selection(self) -> dict:
        app = self._get_app()
        sel = app.selection
        if sel is None:
            raise ExcelError("No cells are currently selected in Excel.")

        raw = sel.value
        if raw is None:
            return {
                "data": [],
                "range": sel.address,
                "sheet": sel.sheet.name,
                "workbook": sel.sheet.book.name,
                "rows": 0,
                "columns": 0,
            }

        # Normalise to 2D list
        if not isinstance(raw, list):
            data = [[raw]]
        elif raw and not isinstance(raw[0], list):
            data = [raw]
        else:
            data = raw

        data = [_serialize_row(row) for row in data]

        return {
            "data": data,
            "range": sel.address,
            "sheet": sel.sheet.name,
            "workbook": sel.sheet.book.name,
            "rows": len(data),
            "columns": len(data[0]) if data else 0,
        }

    def activate_sheet(
        self,
        workbook: str,
        sheet: str,
    ) -> dict:
        wb = self._get_workbook(workbook)
        ws = self._get_sheet(wb, sheet)
        ws.activate()
        return {
            "message": f"Sheet '{sheet}' is now active.",
            "workbook": wb.name,
            "active_sheet": ws.name,
        }

    def run_macro(
        self,
        macro_name: str,
        workbook: str | None = None,
        args: list | None = None,
    ) -> dict:
        app = self._get_app()

        # Build the fully-qualified macro name if workbook is specified
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

    def close_workbook(
        self,
        workbook: str,
        save: bool = False,
    ) -> dict:
        wb = self._get_workbook(workbook)
        name = wb.name
        if save:
            wb.save()
        wb.close()
        return {"message": f"Workbook '{name}' closed (save={save})."}

    def recalculate(self, workbook: str | None = None) -> dict:
        app = self._get_app()
        if workbook:
            wb = self._get_workbook(workbook)
            wb.app.calculate()
            return {"message": f"Workbook '{wb.name}' recalculated."}
        app.calculate()
        return {"message": "All open workbooks recalculated."}

    # ------------------------------------------------------------------ #
    #  Structure manipulation
    # ------------------------------------------------------------------ #

    def insert_delete_cells(
        self,
        workbook: str,
        action: str,
        sheet: str | None = None,
        position: int = 1,
        count: int = 1,
        target: str = "row",
    ) -> dict:
        wb = self._get_workbook(workbook)
        ws = self._get_sheet(wb, sheet)

        if target not in ("row", "column"):
            raise ExcelError("target must be 'row' or 'column'.")
        if count < 1:
            raise ExcelError("count must be >= 1.")

        if action == "insert":
            if target == "row":
                rng_str = f"{position}:{position + count - 1}"
                ws.range(rng_str).api.EntireRow.Insert()
                return {
                    "message": f"Inserted {count} row(s) at row {position}.",
                    "sheet": ws.name,
                }
            else:
                rng = ws.range((1, position), (1, position + count - 1))
                rng.api.EntireColumn.Insert()
                return {
                    "message": f"Inserted {count} column(s) at column {position}.",
                    "sheet": ws.name,
                }

        if action == "delete":
            if target == "row":
                rng_str = f"{position}:{position + count - 1}"
                ws.range(rng_str).api.EntireRow.Delete()
                return {
                    "message": f"Deleted {count} row(s) starting at row {position}.",
                    "sheet": ws.name,
                }
            else:
                rng = ws.range((1, position), (1, position + count - 1))
                rng.api.EntireColumn.Delete()
                return {
                    "message": f"Deleted {count} column(s) starting at column {position}.",
                    "sheet": ws.name,
                }

        raise ExcelError(
            f"Unknown action '{action}'. Use: insert, delete."
        )
