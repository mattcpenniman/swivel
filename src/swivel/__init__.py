"""swivel - add refreshed pivot tables to Excel .xlsx files.

    import swivel
    swivel.add_pivot("sales.xlsx", source="Data", target="Summary",
                     rows=["Region"], columns=["Year"], values=["Amount"])

Every pivot table swivel writes is *already refreshed* - the cache, its
records and the rendered cells are computed from the current data - and is
flagged to refresh again whenever Excel opens the file.
"""
__version__ = "0.1.0"

from ._values import XLError  # noqa: E402
from .api import (PivotResult, add_pivot, add_pivots, list_pivots,  # noqa: E402
                  set_refresh_on_open)
from .spec import ALL, PivotSpec, SwivelError, Value  # noqa: E402

__all__ = ["add_pivot", "add_pivots", "list_pivots", "set_refresh_on_open", "PivotSpec",
           "Value", "PivotResult", "SwivelError", "XLError", "ALL", "__version__"]
