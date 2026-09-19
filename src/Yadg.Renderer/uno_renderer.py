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


def refresh(document):
    fields = document.getTextFields()
    if hasattr(fields, "refresh"):
        fields.refresh()
    indexes = document.getDocumentIndexes()
    for index in indexes:
        index.update()
    if hasattr(document, "updateLinks"):
        document.updateLinks()
    if hasattr(document, "calculateAll"):
        document.calculateAll()
    if hasattr(fields, "refresh"):
        fields.refresh()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--docx-output", required=True)
    parser.add_argument("--pdf-output", required=True)
    args = parser.parse_args()
    document = None
    try:
        context = connect(args.port)
        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        document = desktop.loadComponentFromURL(file_url(args.input), "_blank", 0, (prop("Hidden", True), prop("ReadOnly", False), prop("UpdateDocMode", 3)))
        if document is None:
            raise RuntimeError(f"LibreOffice could not open '{args.input}'.")
        refresh(document)
        document.storeToURL(file_url(args.docx_output), (prop("FilterName", "Office Open XML Text"), prop("Overwrite", True)))
        document.storeToURL(file_url(args.pdf_output), (prop("FilterName", "writer_pdf_Export"), prop("Overwrite", True)))
        print(json.dumps({"ok": True}), flush=True)
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}), flush=True)
        return 1
    finally:
        if document is not None:
            try:
                document.close(True)
            except Exception:
                try:
                    document.dispose()
                except Exception:
                    pass


if __name__ == "__main__":
    sys.exit(main())
