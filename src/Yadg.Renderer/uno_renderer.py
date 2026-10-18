import argparse
import json
import os
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def file_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(path))


def report_progress(phase, status, started, phase_started=None):
    now = time.monotonic()
    event = {
        "type": "progress",
        "phase": phase,
        "status": status,
        "elapsedMs": (now - started) * 1000,
    }
    if phase_started is not None:
        event["phaseMs"] = (now - phase_started) * 1000
    print(json.dumps(event), flush=True)


def run_phase(name, started, action):
    phase_started = time.monotonic()
    report_progress(name, "started", started)
    result = action()
    report_progress(name, "complete", started, phase_started)
    return result


def connect(port):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    last = None
    for _ in range(60):
        try:
            return resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
        except Exception as exc:
            last = exc
            time.sleep(0.25)
    raise RuntimeError(f"Could not connect to isolated LibreOffice UNO session: {last}")


def refresh_fields(document):
    fields = document.getTextFields()
    if hasattr(fields, "refresh"):
        fields.refresh()


def update_indexes(document):
    indexes = document.getDocumentIndexes()
    for index in indexes:
        index.update()


def update_links(document):
    if hasattr(document, "updateLinks"):
        document.updateLinks()


def calculate(document):
    if hasattr(document, "calculateAll"):
        document.calculateAll()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--docx-output", required=True)
    parser.add_argument("--pdf-output", required=True)
    args = parser.parse_args()
    document = None
    started = time.monotonic()
    current_phase = "UNO connect"
    result = {"type": "result", "ok": True}
    exit_code = 0
    try:
        context = run_phase(current_phase, started, lambda: connect(args.port))
        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        current_phase = "document load"
        document = run_phase(current_phase, started, lambda: desktop.loadComponentFromURL(file_url(args.input), "_blank", 0, (prop("Hidden", True), prop("ReadOnly", False), prop("UpdateDocMode", 3))))
        if document is None:
            raise RuntimeError(f"LibreOffice could not open '{args.input}'.")
        current_phase = "field refresh"
        run_phase(current_phase, started, lambda: refresh_fields(document))
        current_phase = "index update"
        run_phase(current_phase, started, lambda: update_indexes(document))
        current_phase = "link update"
        run_phase(current_phase, started, lambda: update_links(document))
        current_phase = "calculation"
        run_phase(current_phase, started, lambda: calculate(document))
        current_phase = "field refresh"
        run_phase(current_phase, started, lambda: refresh_fields(document))
        current_phase = "DOCX save"
        run_phase(current_phase, started, lambda: document.storeToURL(file_url(args.docx_output), (prop("FilterName", "Office Open XML Text"), prop("Overwrite", True))))
        current_phase = "PDF export"
        run_phase(current_phase, started, lambda: document.storeToURL(file_url(args.pdf_output), (prop("FilterName", "writer_pdf_Export"), prop("Overwrite", True))))
    except Exception as exc:
        result = {"type": "result", "ok": False, "phase": current_phase, "error": str(exc)}
        exit_code = 1
    finally:
        if document is not None:
            try:
                current_phase = "document close"
                run_phase(current_phase, started, lambda: document.close(True))
            except Exception:
                try:
                    document.dispose()
                except Exception:
                    pass
    print(json.dumps(result), flush=True)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
