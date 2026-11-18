"""Optional: drive a real Excel (Windows + ``pip install swivel[excel]``).

swivel's files are already refreshed when written and refresh again when
opened, so Excel is never required. These helpers exist for two jobs:

* :func:`refresh` - have Excel itself recompute and save every pivot table in
  a workbook (useful for workbooks whose pivot tables swivel did not create);
* :func:`snapshot` - read back what Excel shows, which the test-suite uses as
  an oracle.
"""
from __future__ import annotations

import contextlib
import os
from dataclasses import dataclass

# Excel's COM error codes for cell error values
COM_ERRORS = {
    -2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?",
    -2146826288: "#NULL!", -2146826252: "#NUM!", -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


@contextlib.contextmanager
def excel_app():
    """A private, invisible Excel instance that is always quit afterwards."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    try:
        # early binding: late-bound COM sometimes mistakes methods for properties
        app = win32com.client.gencache.EnsureDispatch(app)
    except Exception:
        pass
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.EnableEvents = False
        yield app
    finally:
        with contextlib.suppress(Exception):
            app.DisplayAlerts = False
            for wb in list(app.Workbooks):
                wb.Close(False)
        with contextlib.suppress(Exception):
            app.Quit()
        del app
        pythoncom.CoUninitialize()


_BUSY = (-2147418111, -2147417846)  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


def _retry(fn, timeout: float = 120.0):
    """Run ``fn`` again while Excel says it is busy (e.g. refreshing on open)."""
    import time
    import pywintypes
    deadline = time.monotonic() + timeout
    while True:
        try:
            return fn()
        except pywintypes.com_error as exc:
            if exc.args[0] not in _BUSY or time.monotonic() > deadline:
                raise
            time.sleep(0.25)


def _open(xl, path, read_only):
    wb = _retry(lambda: xl.Workbooks.Open(path, 0, read_only))
    _retry(lambda: xl.CalculateUntilAsyncQueriesDone())
    return wb


def _close(wb):
    with contextlib.suppress(Exception):
        _retry(lambda: wb.Close(False), timeout=30)


def refresh(path, *, output=None, app=None) -> int:
    """Open ``path`` in Excel, refresh every pivot table, save. Returns #pivots."""
    path = os.path.abspath(os.fspath(path))
    out = os.path.abspath(os.fspath(output)) if output else path
    ctx = contextlib.nullcontext(app) if app is not None else excel_app()
    with ctx as xl:
        wb = _open(xl, path, False)
        try:
            def work():
                n = 0
                for ws in _sheets(wb):
                    for pt in _pivot_tables(ws):
                        pt.RefreshTable()
                        n += 1
                if out == path:
                    wb.Save()
                else:
                    wb.SaveAs(out, wb.FileFormat)
                return n
            return _retry(work)
        finally:
            _close(wb)


@dataclass
class PivotSnapshot:
    sheet: str
    name: str
    table_range1: str
    table_range2: str
    cells: dict          # {"A3": value}; errors as "#N/A" strings, dates as serials


def _sheets(wb):
    for i in range(1, wb.Worksheets.Count + 1):
        yield wb.Worksheets(i)


def _pivot_tables(ws):
    pts = ws.PivotTables()
    for i in range(1, pts.Count + 1):
        yield ws.PivotTables(i)


def _norm(v):
    if isinstance(v, int) and v in COM_ERRORS:
        return COM_ERRORS[v]
    return v


def _read_pivots(wb, refresh_first):
    from ._xml import cell_ref
    out = []
    for ws in _sheets(wb):
        for pt in _pivot_tables(ws):
            if refresh_first:
                pt.RefreshTable()
            rng = pt.TableRange2
            vals = rng.Value2
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            r0, c0 = rng.Row - 1, rng.Column - 1
            cells = {}
            for i, row in enumerate(vals):
                for j, v in enumerate(row):
                    if v is not None and v != "":
                        cells[cell_ref(r0 + i, c0 + j)] = _norm(v)
            out.append(PivotSnapshot(ws.Name, pt.Name,
                                     pt.TableRange1.GetAddress(False, False),
                                     rng.GetAddress(False, False), cells))
    return out


def snapshot(path, *, app=None, refresh_first: bool = True) -> list[PivotSnapshot]:
    """Every pivot table as Excel renders it (after an explicit refresh)."""
    path = os.path.abspath(os.fspath(path))
    ctx = contextlib.nullcontext(app) if app is not None else excel_app()
    with ctx as xl:
        wb = _open(xl, path, True)
        try:
            return _retry(lambda: _read_pivots(wb, refresh_first))
        finally:
            _close(wb)
