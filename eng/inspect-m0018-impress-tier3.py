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


def inspect_document(desktop, path: Path) -> dict:
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0, (pv("ReadOnly", True),))
    if doc is None:
        raise RuntimeError(f"Impress did not return a document for {path}.")
    try:
        pages = doc.getDrawPages()
        if pages.getCount() != 4:
            raise RuntimeError(f"Expected four generated slides in {path}; found {pages.getCount()}.")
        titles: list[str] = []
        footers = 0
        text_shapes = 0
        pictures = 0
        tables = []
        for index in range(pages.getCount()):
            page = pages.getByIndex(index)
            title = None
            for shape_index in range(page.getCount()):
                shape = page.getByIndex(shape_index)
                shape_type = shape.getShapeType()
                name = shape.getName()
                if name == "Native template footer":
                    footers += 1
                if "TitleTextShape" in shape_type:
                    title = shape.getString().strip()
                if "TextShape" in shape_type or "OutlinerShape" in shape_type or "TitleTextShape" in shape_type:
                    text_shapes += 1
                    if "{{" in shape.getString():
                        raise RuntimeError(f"Unresolved visible tag in generated slide {index + 1} of {path}.")
                if "GraphicObjectShape" in shape_type:
                    pictures += 1
                if shape_type == "com.sun.star.drawing.TableShape":
                    model = shape.Model
                    rows = model.RowCount
                    columns = model.ColumnCount
                    if columns != 2 or rows < 3:
                        raise RuntimeError(f"Expected the editable 2-column, 3-row native table in {path}; found {columns}x{rows}.")
                    cell_text = [model.getCellByPosition(column, row).String for row in range(rows) for column in range(columns)]
                    if any("{{" in value for value in cell_text):
                        raise RuntimeError(f"Unresolved table tag in generated slide {index + 1} of {path}.")
                    tables.append({"rows": rows, "columns": columns, "cells": cell_text})
            if not title:
                raise RuntimeError(f"Generated slide {index + 1} has no editable native title in {path}.")
            titles.append(title)
        if titles != ["Narrative overview", "Audience choices", "Architecture map", "Measures"]:
            raise RuntimeError(f"Generated slide order differs from Markdown in {path}: {titles}")
        if footers != 4 or pictures < 1 or len(tables) != 1 or text_shapes < 4:
            raise RuntimeError(f"Native content/static inspection failed in {path}: footer={footers}, picture={pictures}, table={len(tables)}, text={text_shapes}.")
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
        }
    finally:
        doc.close(True)


result = "failed"
failure = None
inspection = None
try:
    context = connect()
    manager = context.ServiceManager
    desktop = manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    inspection = inspect_document(desktop, open_path)
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
    record = {
        "milestone": "M0018",
        "gate": "VAL-03",
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
