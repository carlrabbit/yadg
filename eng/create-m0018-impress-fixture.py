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


def set_position_size(shape, x: int, y: int, width: int, height: int) -> None:
    shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", x, y))
    shape.setSize(uno.createUnoStruct("com.sun.star.awt.Size", width, height))


def copy_native_text_format(source, target) -> None:
    for property_name in (
        "CharFontName", "CharHeight", "CharColor", "CharWeight", "CharPosture",
        "ParaAdjust", "ParaLeftMargin", "ParaFirstLineIndent",
    ):
        try:
            target.setPropertyValue(property_name, source.getPropertyValue(property_name))
        except Exception as error:  # noqa: BLE001 - native exemplar transfer must not silently fail
            raise RuntimeError(f"Could not transfer native Impress {property_name} formatting.") from error


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


def fill_example(shape, marker: str, title_exemplar, *, include_heading: bool, include_list: bool) -> None:
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
        copy_native_text_format(title_exemplar, cursor)
    if include_list:
        cursor.gotoStart(False)
        cursor.goRight(len("\n".join(lines[:-1])) + 1, False)
        cursor.NumberingLevel = 0


def set_title(page, layout_id: str):
    for index in range(page.getCount()):
        shape = page.getByIndex(index)
        if "TitleTextShape" in shape.getShapeType():
            shape.Name = f"YADG title {layout_id}"
            shape.String = "{{title}}"
            return shape
    raise RuntimeError(f"Native title placeholder is missing for {layout_id}.")


def add_slide(doc, pages, layout: int, layout_id: str):
    page = pages.insertNewByIndex(pages.getCount())
    page.Layout = layout
    title = set_title(page, layout_id)
    add_marker(doc, page, layout_id)
    add_footer(doc, page)
    return page, title


def get_native_content_placeholders(page):
    return [
        page.getByIndex(index)
        for index in range(page.getCount())
        if "OutlinerShape" in page.getByIndex(index).getShapeType()
    ]


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
        content_title = next(content.getByIndex(index) for index in range(content.getCount()) if "TitleTextShape" in content.getByIndex(index).getShapeType())
        content_region = get_native_content_placeholders(content)
        if len(content_region) != 1:
            raise RuntimeError(f"Expected one native Impress content placeholder, found {len(content_region)}.")
        content_region[0].Name = "YADG content region"
        fill_example(content_region[0], "{{content}}", content_title, include_heading=True, include_list=True)

        columns, columns_title = add_slide(doc, pages, 4, "two-columns")
        outliners = get_native_content_placeholders(columns)
        if len(outliners) != 2:
            raise RuntimeError(f"Expected two native content placeholders, found {len(outliners)}.")
        # The imported Impress "Two Content" master serializes both native
        # outliners with the same full-width geometry. Keep the placeholders
        # and their native text styles, but give each region distinct bounds.
        set_position_size(outliners[0], 1000, 3600, 11800, 9500)
        set_position_size(outliners[1], 14600, 3600, 11800, 9500)
        outliners[0].Name = "YADG left region"
        outliners[1].Name = "YADG right region"
        fill_example(outliners[0], "{{left}}", columns_title, include_heading=True, include_list=True)
        fill_example(outliners[1], "{{right}}", columns_title, include_heading=True, include_list=True)

        visual, visual_title = add_slide(doc, pages, 8, "text-image")
        visual_content = get_native_content_placeholders(visual)
        if len(visual_content) != 1:
            raise RuntimeError(f"Expected one native Impress text-image placeholder, found {len(visual_content)}.")
        visual_region = visual_content[0]
        visual_region.Name = "YADG text region"
        set_position_size(visual_region, 1000, 3600, 10500, 10500)
        fill_example(visual_region, "{{text}}", visual_title, include_heading=True, include_list=True)
        image_region = add_text_shape(doc, visual, "YADG image region", "{{image}}", 12500, 4100, 12000, 10500)
        source_style = visual_region.createTextCursor()
        source_style.gotoStart(False)
        source_style.goRight(len("{{text}}") + 1, False)
        source_style.goRight(len("ordinary paragraph example"), True)
        target_style = image_region.createTextCursor()
        target_style.gotoStart(False)
        target_style.goRight(len("{{image}}"), True)
        copy_native_text_format(source_style, target_style)

        table_page, _ = add_slide(doc, pages, 19, "table")
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
        "creationMethod": "Created as a new presentation through real LibreOffice Impress UNO; native default master, theme, layouts and title/body/list placeholders were retained, heading examples copied native title formatting, a native editable table was authored, and the four prototype slides were saved from Impress as PPTX.",
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
