"""Open and inspect the M0018 generated candidate in real LibreOffice Impress."""

from __future__ import annotations

import hashlib
import json
import os
import platform
import sys
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


template_path, source_path, open_path, candidate_path, evidence_path, repo_path = (Path(value) for value in sys.argv[1:7])
revision = sys.argv[7]


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def pv(name: str, value: object) -> PropertyValue:
    result = PropertyValue()
    result.Name = name
    result.Value = value
    return result


def require_shape(page, name: str):
    for index in range(page.getCount()):
        shape = page.getByIndex(index)
        if shape.getName() == name:
            return shape
    raise RuntimeError(f"Expected native Impress shape '{name}' is missing.")


def require_title_shape(page):
    for index in range(page.getCount()):
        shape = page.getByIndex(index)
        if "TitleTextShape" in shape.getShapeType():
            return shape
    raise RuntimeError("Expected native Impress title placeholder is missing.")


def require_text_shape(page, sample: str):
    for index in range(page.getCount()):
        shape = page.getByIndex(index)
        if sample in shape.getString():
            return shape
    raise RuntimeError(f"Expected native Impress text region containing '{sample}' is missing.")


def select_text(shape, value: str):
    position = shape.getString().find(value)
    if position < 0:
        raise RuntimeError(f"Shape '{shape.getName()}' does not contain formatting sample '{value}'.")
    cursor = shape.createTextCursor()
    cursor.gotoStart(False)
    if position:
        cursor.goRight(position, False)
    cursor.goRight(len(value), True)
    return cursor


def formatting_signature(text_range) -> dict:
    names = (
        "CharFontName", "CharHeight", "CharColor", "CharWeight", "CharPosture",
        "ParaAdjust", "ParaLeftMargin", "ParaFirstLineIndent", "NumberingLevel",
    )
    signature = {}
    for name in names:
        try:
            value = text_range.getPropertyValue(name)
        except Exception as error:  # noqa: BLE001 - a missing style property invalidates this comparison
            raise RuntimeError(f"Native Impress formatting property '{name}' is unavailable.") from error
        if value is None:
            signature[name] = None
            continue
        if name in ("CharFontName", "CharPosture", "ParaAdjust"):
            signature[name] = str(value)
        elif name == "CharHeight":
            signature[name] = round(float(value), 3)
        else:
            signature[name] = int(value)
    return signature


def compare_formatting(role: str, template_signature: dict, generated_signature: dict) -> dict:
    if template_signature != generated_signature:
        differing = {
            key: {"template": template_signature.get(key), "generated": generated_signature.get(key)}
            for key in template_signature.keys() | generated_signature.keys()
            if template_signature.get(key) != generated_signature.get(key)
        }
        raise RuntimeError(f"Generated {role} formatting differs from its Impress template example: {differing}")
    return {"role": role, "result": "matched-template-example", "template": template_signature, "generated": generated_signature}


def inspect_formatting(desktop, path: Path) -> dict:
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0, (pv("ReadOnly", True),))
    if doc is None:
        raise RuntimeError(f"Impress did not return a document for formatting inspection: {path}.")
    try:
        pages = doc.getDrawPages()
        content_page = pages.getByIndex(0)
        table_page = pages.getByIndex(3)
        is_template = path.resolve() == template_path.resolve()
        title = require_title_shape(content_page)
        body_sample = "This paragraph was authored in the shared M0018 compatibility source." if not is_template else "ordinary paragraph example"
        body = require_text_shape(content_page, body_sample)
        table_shape = next(
            table_page.getByIndex(index)
            for index in range(table_page.getCount())
            if table_page.getByIndex(index).getShapeType() == "com.sun.star.drawing.TableShape"
        )
        table = table_shape.Model
        samples = {
            "title": formatting_signature(select_text(title, "{{title}}" if is_template else "Narrative overview")),
            "paragraph": formatting_signature(select_text(body, "ordinary paragraph example" if is_template else "This paragraph was authored in the shared M0018 compatibility source.")),
            "heading": formatting_signature(select_text(body, "content heading example" if is_template else "Design intent")),
            "list": formatting_signature(select_text(body, "list formatting example" if is_template else "Readable default body formatting")),
            "table-header": formatting_signature(table.getCellByPosition(1, 0)),
            "table-body": formatting_signature(table.getCellByPosition(0, 1)),
        }
        return samples
    finally:
        doc.close(True)


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    port = os.environ["YADG_M0018_IMPRESS_PORT"]
    last_error = None
    for _ in range(60):
        try:
            return resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
        except Exception as error:  # noqa: BLE001 - wait for the isolated desktop process
            last_error = error
            time.sleep(0.5)
    raise RuntimeError("Could not connect to the isolated LibreOffice Impress service.") from last_error


def inspect_document(desktop, path: Path, generated: bool) -> dict:
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0, (pv("ReadOnly", True),))
    if doc is None:
        raise RuntimeError(f"Impress did not return a document for {path}.")
    try:
        pages = doc.getDrawPages()
        if pages.getCount() not in (4, 5):
            raise RuntimeError(f"Expected four compatibility slides and an optional recovery slide in {path}; found {pages.getCount()}.")
        titles: list[str] = []
        footers = 0
        text_shapes = 0
        pictures = 0
        tables = []
        layout_relations = []
        for index in range(pages.getCount()):
            page = pages.getByIndex(index)
            title = None
            title_shape = require_title_shape(page)
            prototype_name = title_shape.getString().strip()
            layout_code = int(page.getPropertyValue("Layout"))
            master_name = str(page.getMasterPage().getName())
            placeholder_roles = []
            for shape_index in range(page.getCount()):
                shape = page.getByIndex(shape_index)
                shape_type = shape.getShapeType()
                name = shape.getName()
                if name == "Native template footer":
                    footers += 1
                if "TitleTextShape" in shape_type:
                    title = shape.getString().strip()
                if "TitleTextShape" in shape_type or "OutlinerShape" in shape_type:
                    placeholder_roles.append({"name": name, "shapeType": shape_type})
                if "TextShape" in shape_type or "OutlinerShape" in shape_type or "TitleTextShape" in shape_type:
                    text_shapes += 1
                    if generated and "{{" in shape.getString():
                        raise RuntimeError(f"Unresolved visible tag in generated slide {index + 1} of {path}.")
                if "GraphicObjectShape" in shape_type:
                    pictures += 1
                if shape_type == "com.sun.star.drawing.TableShape":
                    model = shape.Model
                    rows = model.RowCount
                    columns = model.ColumnCount
                    if columns != 2 or rows < (3 if generated else 2):
                        raise RuntimeError(f"Expected the editable 2-column, 3-row native table in {path}; found {columns}x{rows}.")
                    cell_text = [model.getCellByPosition(column, row).String for row in range(rows) for column in range(columns)]
                    if generated and any("{{" in value for value in cell_text):
                        raise RuntimeError(f"Unresolved table tag in generated slide {index + 1} of {path}.")
                    tables.append({"rows": rows, "columns": columns, "cells": cell_text})
            if not title:
                raise RuntimeError(f"Generated slide {index + 1} has no editable native title in {path}.")
            titles.append(title)
            layout_relations.append({
                "slide": index + 1,
                "prototype": prototype_name,
                "layoutCode": layout_code,
                "masterPageName": master_name,
                "placeholderRoles": placeholder_roles,
            })
        expected_titles = ["Narrative overview", "Audience choices", "Architecture map", "Measures"]
        if pages.getCount() == 5:
            expected_titles.append("YOLO recovery")
        if generated and titles != expected_titles:
            raise RuntimeError(f"Generated slide order differs from Markdown in {path}: {titles}")
        if footers != pages.getCount() or (generated and pictures < 1) or len(tables) != 1 or text_shapes < 4:
            raise RuntimeError(f"Native content/static inspection failed in {path}: footer={footers}, picture={pictures}, table={len(tables)}, text={text_shapes}.")
        if generated and pages.getCount() == 5:
            recovery = pages.getByIndex(4)
            recovery_text = []
            for shape_index in range(recovery.getCount()):
                try:
                    recovery_text.append(recovery.getByIndex(shape_index).getString())
                except Exception:
                    pass
            if not any("YADG image unavailable" in value for value in recovery_text):
                raise RuntimeError("The designated YOLO recovery slide has no visible image placeholder.")
        masters = doc.getMasterPages().getCount()
        if masters < 1:
            raise RuntimeError(f"No native Impress master page remains in {path}.")
        if os.environ.get("YADG_M0018_PAUSE_FOR_VISUAL_INSPECTION") == "1":
            print("Impress has opened the generated M0018 candidate. Inspect the visible document; press Enter here after checking that no repair warning is present.", flush=True)
            input()
        return {
            "slideCount": pages.getCount(),
            "slideTitles": titles,
            "staticFooterCount": footers,
            "editablePictureCount": pictures,
            "nativeTableCount": len(tables),
            "nativeTables": tables,
            "editableTextShapeCount": text_shapes,
            "masterPageCount": masters,
            "slideLayoutMasterRelationships": layout_relations,
        }
    finally:
        doc.close(True)


def export_native_pdf(desktop, source: Path, output: Path) -> dict:
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(source)), "_blank", 0, (pv("ReadOnly", True),))
    if doc is None:
        raise RuntimeError(f"Impress could not open a document for native PDF rendering: {source}.")
    try:
        pages = int(doc.getDrawPages().getCount())
        output.parent.mkdir(parents=True, exist_ok=True)
        doc.storeToURL(uno.systemPathToFileUrl(str(output)), (pv("FilterName", "impress_pdf_Export"), pv("Overwrite", True)))
        if not output.is_file() or output.stat().st_size == 0:
            raise RuntimeError(f"Impress did not export its native slide rendering: {output}.")
        return {"path": str(output.relative_to(repo_path)).replace("\\", "/"), "sha256": digest(output), "pageCount": pages, "renderedBy": "LibreOffice Impress native PDF export"}
    finally:
        doc.close(True)


def export_native_slide_images(context, desktop, source: Path, output_dir: Path) -> list[dict]:
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(source)), "_blank", 0, (pv("ReadOnly", True),))
    if doc is None:
        raise RuntimeError(f"Impress could not open a document for native slide rendering: {source}.")
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        pages = doc.getDrawPages()
        records = []
        for index in range(pages.getCount()):
            output = output_dir / f"slide-{index + 1}.png"
            exporter = context.ServiceManager.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", context)
            exporter.setSourceDocument(pages.getByIndex(index))
            exporter.filter((pv("URL", uno.systemPathToFileUrl(str(output))), pv("MediaType", "image/png")))
            if not output.is_file() or output.stat().st_size == 0:
                raise RuntimeError(f"Impress did not export native slide image {index + 1} from {source}.")
            records.append({"slide": index + 1, "path": str(output.relative_to(repo_path)).replace("\\", "/"), "sha256": digest(output), "renderedBy": "LibreOffice Impress GraphicExportFilter"})
        return records
    finally:
        doc.close(True)


result = "failed"
failure = None
inspection = None
try:
    context = connect()
    manager = context.ServiceManager
    desktop = manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    inspection = inspect_document(desktop, open_path, generated=True)
    template_inspection = inspect_document(desktop, template_path, generated=False)
    prototypes = {
        (item["layoutCode"], item["masterPageName"])
        for item in template_inspection["slideLayoutMasterRelationships"]
    }
    for relation in inspection["slideLayoutMasterRelationships"]:
        if (relation["layoutCode"], relation["masterPageName"]) not in prototypes:
            raise RuntimeError(f"Impress output slide {relation['slide']} changed its native slide-layout/master relationship.")
    template_formatting = inspect_formatting(desktop, template_path)
    generated_formatting = inspect_formatting(desktop, open_path)
    if template_formatting["heading"]["CharHeight"] >= template_formatting["title"]["CharHeight"] or generated_formatting["heading"]["CharHeight"] >= generated_formatting["title"]["CharHeight"]:
        raise RuntimeError("An Impress body content heading is at least as large as its slide-title role.")
    formatting_comparisons = [
        compare_formatting(role, template_formatting[role], generated_formatting[role])
        for role in ("title", "paragraph", "heading", "list", "table-header", "table-body")
    ]
    inspection["formattingValidation"] = {
        "result": "passed",
        "source": "LibreOffice Impress UNO formatting comparison against the canonical Impress-origin template",
        "comparisons": formatting_comparisons,
    }
    render_root = evidence_path.parent / "renderings" / "impress-origin"
    native_renderings = {
        "application": "LibreOffice Impress",
        "applicationVersion": os.environ.get("YADG_M0018_IMPRESS_VERSION", ""),
        "prototype": export_native_pdf(desktop, template_path, render_root / "prototype.pdf"),
        "candidate": export_native_pdf(desktop, open_path, render_root / "candidate.pdf"),
        "prototypeSlideImages": export_native_slide_images(context, desktop, template_path, render_root / "prototype"),
        "candidateSlideImages": export_native_slide_images(context, desktop, open_path, render_root / "candidate"),
    }
    if native_renderings["candidate"]["pageCount"] != inspection["slideCount"]:
        raise RuntimeError("Impress native slide rendering does not include every generated slide.")
    inspection["nativeRenderings"] = native_renderings
    desktop.terminate()
    result = "passed"
except Exception as error:  # noqa: BLE001 - leave a durable failure receipt
    failure = f"{type(error).__name__}: {error}"
    try:
        desktop.terminate()
    except Exception:
        pass
finally:
    provenance_path = template_path.with_name("impress-origin.provenance.json")
    provenance = json.loads(provenance_path.read_text(encoding="utf-8")) if provenance_path.exists() else {}
    config_path = Path(os.environ["YADG_M0018_CONFIG"]) if os.environ.get("YADG_M0018_CONFIG") else None
    record = {
        "milestone": os.environ.get("YADG_M0018_MILESTONE", "M0018"),
        "gate": os.environ.get("YADG_M0018_GATE", "VAL-03"),
        "originPath": "Impress-origin template -> YADG -> LibreOffice Impress",
        "result": result,
        "failure": failure,
        "repositoryRevision": revision,
        "application": "LibreOffice Impress",
        "applicationVersion": os.environ.get("YADG_M0018_IMPRESS_VERSION", ""),
        "platform": platform.platform(),
        "templateOriginApplication": "LibreOffice Impress",
        "templateOriginVersion": provenance.get("applicationVersion", ""),
        "templateCreationMethod": provenance.get("creationMethod", ""),
        "templateCreatedAtUtc": provenance.get("createdAtUtc", ""),
        "syntheticAndRedistributionSafe": True,
        "templatePath": str(template_path.relative_to(repo_path)).replace("/", "\\"),
        "templateSha256": digest(template_path),
        "sourcePath": str(source_path.relative_to(repo_path)).replace("/", "\\"),
        "sourceSha256": digest(source_path),
        "configPath": str(config_path.relative_to(repo_path)).replace("/", "\\") if config_path else None,
        "configSha256": digest(config_path) if config_path else None,
        "yolo": os.environ.get("YADG_M0018_YOLO", "False").lower() == "true",
        "outputPath": str(candidate_path.relative_to(repo_path)).replace("/", "\\"),
        "outputSha256": digest(candidate_path) if candidate_path.exists() else None,
        "openedCopyPath": str(open_path),
        "openedCopySha256": digest(open_path) if open_path.exists() else None,
        "openResult": "opened-without-repair" if result == "passed" else "failed",
        "inspection": inspection,
        "profileCleaned": False,
    }
    if evidence_path.exists():
        evidence_path.unlink()
    evidence_path.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")

print(json.dumps({"result": result, "failure": failure, "inspection": inspection}, indent=2))
if result != "passed":
    raise SystemExit(1)
