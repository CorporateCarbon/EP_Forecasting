"""Long-path handling for Excel automation.

Excel's ``Workbooks.Open`` and ``Workbook.SaveAs`` fail once the full path
reaches 256 characters, raising only a bare COM error
(``Open method of Workbooks class failed``, ``0x800A03EC``) that says nothing
about the real cause. Measured on Excel 16.0: 255 characters opens, 256 fails.

CCG's SharePoint project trees routinely exceed this, e.g.::

    …\\08. ERF Projects\\Blackwood Biodiversity and Carbon Project - DP
    \\01. Assessment & Forecasts\\2. FY26 Forecasts\\FY26 review files
    \\260803_BLK_Calculator_CEA01-12_No treat.xlsx        (258 characters)

The fix is to hand Excel a Windows 8.3 short name, which it accepts and
resolves to the same file, so the workbook stays where it is and nothing is
copied or moved. ``xlwings`` cannot be used directly for this: both
``Books.open`` and ``Book.save`` call ``os.path.realpath`` internally, which
expands 8.3 names straight back to the long form.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict

import win32api
import xlwings as xw

EXCEL_MAX_PATH = 255

# Excel FileFormat constants, mirroring xlwings' own extension mapping.
EXT_FILE_FORMAT: Dict[str, int] = {
    ".xlsx": 51,  # xlOpenXMLWorkbook
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12
    ".xls": 56,  # xlExcel8
}


def shorten_for_excel(path: Path) -> str:
    """Return a path string short enough for Excel to open or save.

    Args:
        path: Target file. It may or may not exist yet; if it does not, its
            parent directory must.

    Returns:
        A path string of at most ``EXCEL_MAX_PATH`` characters.

    Raises:
        ValueError: If the path is still too long after shortening, which
            happens when 8.3 name generation is disabled on the volume.
    """
    text = str(path)
    if len(text) <= EXCEL_MAX_PATH:
        return text

    if path.exists():
        short = win32api.GetShortPathName(text)
    else:
        # New file: only the existing parent can be shortened.
        short = str(Path(win32api.GetShortPathName(str(path.parent))) / path.name)

    if len(short) > EXCEL_MAX_PATH:
        raise ValueError(
            f"Excel cannot handle paths longer than {EXCEL_MAX_PATH} characters, "
            f"and this one is {len(text)} even after shortening to {len(short)}:\n"
            f"{text}\n"
            "Move or save the file to a folder with a shorter path."
        )
    return short


def open_workbook(app: xw.App, path: str | Path) -> xw.Book:
    """Open a workbook, tolerating paths beyond Excel's length limit.

    Drop-in replacement for ``app.books.open(path)``.

    Args:
        app: Running Excel application.
        path: Workbook to open.

    Returns:
        The opened workbook.

    Raises:
        FileNotFoundError: If the workbook does not exist.
    """
    target = Path(path)
    if len(str(target)) <= EXCEL_MAX_PATH:
        return app.books.open(str(target))

    if not target.exists():
        raise FileNotFoundError(f"No such file: '{target}'")

    com_book = app.api.Workbooks.Open(shorten_for_excel(target))
    return app.books[com_book.Name]


def save_workbook(book: xw.Book, path: str | Path) -> None:
    """Save a workbook to ``path``, tolerating paths beyond Excel's limit.

    Drop-in replacement for ``book.save(path)``. Saving in place with no path
    (``book.save()``) needs no equivalent, as Excel keeps the handle it already
    opened the file with.

    Args:
        book: Workbook to save.
        path: Destination. Its parent directory must already exist.

    Raises:
        ValueError: If an over-length path carries an extension Excel's
            ``SaveAs`` cannot be given a file format for.
    """
    target = Path(path)
    if len(str(target)) <= EXCEL_MAX_PATH:
        book.save(str(target))
        return

    suffix = target.suffix.lower()
    if suffix not in EXT_FILE_FORMAT:
        raise ValueError(
            f"Cannot save '{suffix}' to a path longer than {EXCEL_MAX_PATH} "
            f"characters. Choose a shorter output path, or save as one of: "
            f"{', '.join(sorted(EXT_FILE_FORMAT))}."
        )

    with book.app.properties(display_alerts=False):
        book.api.SaveAs(
            shorten_for_excel(target), FileFormat=EXT_FILE_FORMAT[suffix]
        )
