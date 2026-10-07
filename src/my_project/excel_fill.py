"""Fill rows of a downloaded .xlsx template by patching the sheet XML directly.

openpyxl drops the template's dropdown validations on save, so cells are written
in place and every other part of the workbook is copied through untouched.
"""
import copy
import datetime as dt
import re
import zipfile

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ROW, CELL = f"{{{NS}}}row", f"{{{NS}}}c"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
EPOCH = dt.date(1899, 12, 30)


def _col_index(ref: str) -> int:
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + ord(ch) - 64
    return n


def _sheet_part(zf: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = etree.fromstring(zf.read("xl/workbook.xml"))
    rid = next(
        (s.get(f"{{{RNS}}}id") for s in workbook.iter(f"{{{NS}}}sheet") if s.get("name") == sheet_name),
        None,
    )
    if rid is None:
        raise KeyError(f"Sheet {sheet_name!r} not found in template")
    rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return target.lstrip("/") if target.startswith("/") else f"xl/{target}"


def _blank(row: etree._Element, number: int) -> None:
    row.set("r", str(number))
    for cell in row.iter(CELL):
        cell.set("r", re.match(r"[A-Z]+", cell.get("r")).group() + str(number))
        cell.attrib.pop("t", None)
        for child in list(cell):
            cell.remove(child)


def _get_row(sheet_data: etree._Element, number: int, style_row: int) -> etree._Element:
    insert_at = len(sheet_data)
    for i, row in enumerate(sheet_data):
        n = int(row.get("r"))
        if n == number:
            return row
        if n > number:
            insert_at = i
            break
    model = next((r for r in sheet_data if r.get("r") == str(style_row)), None)
    new = copy.deepcopy(model) if model is not None else etree.Element(ROW)
    _blank(new, number)
    sheet_data.insert(insert_at, new)
    return new


def _set_cell(row: etree._Element, ref: str, value) -> None:
    idx = _col_index(ref)
    cell, insert_at = None, len(row)
    for i, c in enumerate(row):
        ci = _col_index(c.get("r"))
        if ci == idx:
            cell = c
            break
        if ci > idx:
            insert_at = i
            break
    if cell is None:
        cell = etree.Element(CELL)
        cell.set("r", ref)
        row.insert(insert_at, cell)
    cell.attrib.pop("t", None)
    for child in list(cell):
        cell.remove(child)
    if value is None or value == "":
        return
    if isinstance(value, dt.datetime):
        value = value.date()
    if isinstance(value, dt.date):
        etree.SubElement(cell, f"{{{NS}}}v").text = str((value - EPOCH).days)
    elif isinstance(value, (int, float)):
        etree.SubElement(cell, f"{{{NS}}}v").text = str(value)
    else:
        cell.set("t", "inlineStr")
        text = etree.SubElement(etree.SubElement(cell, f"{{{NS}}}is"), f"{{{NS}}}t")
        text.text = str(value).replace("\xa0", " ").strip()
        text.set(XML_SPACE, "preserve")


def _copy_style(model_row: etree._Element, row: etree._Element, col: str, number: int) -> None:
    style = next((c.get("s") for c in model_row if c.get("r") == f"{col}{model_row.get('r')}"), None)
    cell = next(c for c in row if c.get("r") == f"{col}{number}")
    if style is not None:
        cell.set("s", style)


def fill_template(src, dst, records, sheet="Onboarding", first_row=10) -> None:
    """Write one record per row, starting at first_row. Each record maps column letters to values.

    Rows past the template's styled rows copy the style of first_row.
    """
    with zipfile.ZipFile(src) as zin:
        part = _sheet_part(zin, sheet)
        root = etree.fromstring(zin.read(part))
        sheet_data = root.find(f"{{{NS}}}sheetData")
        for offset, record in enumerate(records):
            number = first_row + offset
            row = _get_row(sheet_data, number, first_row)
            model = _get_row(sheet_data, first_row, first_row)
            for col, value in record.items():
                _set_cell(row, f"{col}{number}", value)
                if offset:
                    _copy_style(model, row, col, number)
        patched = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, patched if item.filename == part else zin.read(item.filename))
