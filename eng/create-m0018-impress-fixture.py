"""Create the canonical M0018 Impress-origin fixture through Impress UNO."""

from __future__ import annotations

import hashlib
import json
import os
import platform
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


REPO = Path(__file__).resolve().parent.parent
FIXTURE_DIR = REPO / "tests" / "fixtures" / "m0018"
OUTPUT = FIXTURE_DIR / "impress-origin.pptx"
PROVENANCE = FIXTURE_DIR / "impress-origin.provenance.json"
PORT = os.environ.get("YADG_M0018_IMPRESS_PORT", "20838")


def property_value(name: str, value: object) -> PropertyValue:
    result = PropertyValue()
    result.Name = name
    result.Value = value
    return result


def set_shape_text(shape, text: str) -> None:
    shape.String = text


def add_text_shape(doc, page, name: str, text: str, x: int, y: int, width: int, height: int):
    shape = doc.createInstance("com.sun.star.drawing.TextShape")
    page.add(shape)
    shape.Name = name
    shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", x, y))
    shape.setSize(uno.createUnoStruct("com.sun.star.awt.Size", width, height))
    set_shape_text(shape, text)
    shape.FillStyle = 0
    shape.LineStyle = 0
    return shape


def replace_placeholder_with_text_shape(doc, page, placeholder, name: str, x: int, y: int, width: int, height: int, style_source=None):
    """Materialize inherited body text as an ordinary, template-positioned text shape."""
    shape = doc.createInstance("com.sun.star.drawing.TextShape")
    page.add(shape)
    shape.Name = name
    style_source = style_source or placeholder
    try:
        shape.Style = style_source.Style
    except Exception:  # noqa: BLE001 - some native layout styles are not assignable to ordinary shapes
        pass
    for property_name in (
        "CharFontName", "CharHeight", "CharColor", "CharWeight", "CharPosture",
        "ParaAdjust", "ParaTopMargin", "ParaBottomMargin", "ParaLeftMargin",
        "ParaRightMargin", "TextWordWrap", "TextAutoGrowHeight", "TextAutoGrowWidth",
    ):
        try:
            shape.setPropertyValue(property_name, style_source.getPropertyValue(property_name))
        except Exception:  # noqa: BLE001 - property availability depends on the native placeholder type
            pass
    shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", x, y))
    shape.setSize(uno.createUnoStruct("com.sun.star.awt.Size", width, height))
    shape.FillStyle = 0
    shape.LineStyle = 0
    page.remove(placeholder)
    return shape


def add_marker(doc, page, layout: str) -> None:
    add_text_shape(doc, page, f"YADG prototype {layout}", f"{{{{slide:{layout}}}}}", 24500, 100, 2500, 350)


def add_footer(doc, page) -> None:
    shape = add_text_shape(doc, page, "Native template footer", "M0018 compatibility fixture · Impress native template", 1000, 14900, 21000, 450)
    shape.CharHeight = 9
    shape.CharPosture = uno.Enum("com.sun.star.awt.FontSlant", "ITALIC")


def get_placeholders(page):
    result = {}
    for index in range(page.getCount()):
        shape = page.getByIndex(index)
        shape_type = shape.getShapeType()
        if "TitleTextShape" in shape_type:
            result.setdefault("title", shape)
        elif "OutlinerShape" in shape_type:
            result.setdefault("content", shape)
            result.setdefault(f"content-{len([key for key in result if key.startswith('content-')]) + 1}", shape)
    return result


def fill_example(shape, marker: str, *, include_heading: bool, include_list: bool) -> None:
    lines = [marker, "ordinary paragraph example"]
    if include_heading:
        lines.append("content heading example")
    if include_list:
        lines.append("list formatting example")
    shape.String = "\n".join(lines)
    # Preserve the application-provided placeholder paragraph style. Add role
    # distinctions only as examples, leaving YADG to inherit them.
    cursor = shape.createTextCursor()
    cursor.gotoStart(False)
    if include_heading:
        cursor.goRight(len(marker) + 1 + len("ordinary paragraph example") + 1, False)
        cursor.goRight(len("content heading example"), True)
        cursor.CharHeight = 26
    if include_list:
        cursor.gotoStart(False)
        cursor.goRight(len("\n".join(lines[:-1])) + 1, False)
        cursor.NumberingLevel = 0


def set_title(page, layout_id: str) -> None:
    for index in range(page.getCount()):
        shape = page.getByIndex(index)
        if "TitleTextShape" in shape.getShapeType():
            shape.Name = f"YADG title {layout_id}"
            shape.String = "{{title}}"
            return
    raise RuntimeError(f"Native title placeholder is missing for {layout_id}.")


def add_slide(doc, pages, layout: int, layout_id: str):
    page = pages.insertNewByIndex(pages.getCount())
    page.Layout = layout
    set_title(page, layout_id)
    add_marker(doc, page, layout_id)
    add_footer(doc, page)
    return page


def connect():
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    last_error = None
    for _ in range(60):
        try:
            return resolver.resolve(
                f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext"
            )
        except Exception as error:  # noqa: BLE001 - retry until the isolated app is ready
            last_error = error
            time.sleep(0.5)
    raise RuntimeError("Could not connect to the running LibreOffice Impress UNO service.") from last_error


def main() -> None:
    if OUTPUT.exists() or PROVENANCE.exists():
        raise FileExistsError("Refusing to overwrite an existing canonical Impress fixture.")
    FIXTURE_DIR.mkdir(parents=True, exist_ok=True)
    context = connect()
    manager = context.ServiceManager
    desktop = manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    doc = desktop.loadComponentFromURL("private:factory/simpress", "_blank", 0, ())
    if doc is None:
        raise RuntimeError("Impress did not create a new presentation document.")

    try:
        pages = doc.getDrawPages()
        if pages.getCount() != 1:
            raise RuntimeError("A new Impress presentation should start with one native slide.")
        content = pages.getByIndex(0)
        content.Layout = 8
        set_title(content, "content")
        add_marker(doc, content, "content")
        add_footer(doc, content)
        fill_example(get_placeholders(content)["content"], "{{content}}", include_heading=True, include_list=True)

        columns = add_slide(doc, pages, 4, "two-columns")
        outliners = [columns.getByIndex(i) for i in range(columns.getCount()) if "OutlinerShape" in columns.getByIndex(i).getShapeType()]
        if len(outliners) != 2:
            raise RuntimeError(f"Expected two native content placeholders, found {len(outliners)}.")
        left_region = replace_placeholder_with_text_shape(doc, columns, outliners[0], "YADG left region", 1000, 3600, 12000, 10000)
        right_region = replace_placeholder_with_text_shape(doc, columns, outliners[1], "YADG right region", 14500, 3600, 12000, 10000, left_region)
        fill_example(left_region, "{{left}}", include_heading=True, include_list=True)
        fill_example(right_region, "{{right}}", include_heading=True, include_list=True)

        visual = add_slide(doc, pages, 8, "text-image")
        visual_content = next(visual.getByIndex(i) for i in range(visual.getCount()) if "OutlinerShape" in visual.getByIndex(i).getShapeType())
        visual_region = replace_placeholder_with_text_shape(doc, visual, visual_content, "YADG text region", 1000, 3600, 10500, 10500)
        fill_example(visual_region, "{{text}}", include_heading=True, include_list=True)
        add_text_shape(doc, visual, "YADG image region", "{{image}}", 12500, 4100, 12000, 10500)

        table_page = add_slide(doc, pages, 19, "table")
        doc.getCurrentController().setCurrentPage(table_page)
        frame = doc.getCurrentController().getFrame()
        dispatch = manager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", context)
        dispatch.executeDispatch(frame, ".uno:InsertTable", "", 0, (
            property_value("Rows", 2), property_value("Columns", 2)
        ))
        table_shape = next(
            table_page.getByIndex(i)
            for i in range(table_page.getCount())
            if table_page.getByIndex(i).getShapeType() == "com.sun.star.drawing.TableShape"
        )
        table_shape.Name = "YADG native editable table"
        table_shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", 1000, 3900))
        table_shape.setSize(uno.createUnoStruct("com.sun.star.awt.Size", 24000, 9500))
        table = table_shape.Model
        values = (("{{table}}", "header formatting example"), ("body formatting example", "body formatting example"))
        for row in range(2):
            for column in range(2):
                cell = table.getCellByPosition(column, row)
                cell.String = values[row][column]
                cell.CharHeight = 16
                if row == 0:
                    cell.CharWeight = 150.0
                else:
                    cell.CharPosture = uno.Enum("com.sun.star.awt.FontSlant", "ITALIC")
        rows = table.getRows()
        for row_index in range(rows.getCount()):
            row_model = rows.getByIndex(row_index)
            row_model.OptimalHeight = False
            row_model.Height = 1800
        now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
        output_url = uno.systemPathToFileUrl(str(OUTPUT))
        doc.storeAsURL(output_url, (
            property_value("FilterName", "Impress MS PowerPoint 2007 XML"),
            property_value("Overwrite", False),
        ))
        doc.close(True)
    finally:
        try:
            doc.dispose()
        except Exception:
            pass

    version = os.environ.get("YADG_M0018_IMPRESS_VERSION", "").strip()
    if not version:
        raise RuntimeError("Run eng/create-m0018-impress-fixture.ps1 to supply the installed Impress version.")
    digest = hashlib.sha256(OUTPUT.read_bytes()).hexdigest().upper()
    provenance = {
        "id": "impress-origin",
        "path": str(OUTPUT.relative_to(REPO)).replace("/", "\\"),
        "originApplication": "LibreOffice Impress",
        "applicationVersion": version,
        "platform": platform.platform(),
        "createdAtUtc": now,
        "saveDateUtc": now,
        "creationMethod": "Created as a new presentation through real LibreOffice Impress UNO; native default master, theme, layouts, title/body/list placeholders and a native editable table were authored in the Impress document model before saving as PPTX.",
        "syntheticAndRedistributionSafe": True,
        "createdInApplication": True,
        "notOpenXmlGeneratedOrRoundTripped": True,
        "retainedNativeMasterThemeLayouts": True,
        "layouts": ["content", "two-columns", "text-image", "table"],
        "sha256": digest,
    }
    PROVENANCE.write_text(json.dumps(provenance, indent=2) + "\n", encoding="utf-8")
    print(f"Created Impress-origin fixture: {OUTPUT}")
    print(f"Provenance: {PROVENANCE}")


if __name__ == "__main__":
    main()
