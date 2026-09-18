"""Tag an .xlsx so Excel opens the Excel Assistant task pane with it.

The parts mirror the ExcelWorkbookWithTaskPane.xlsx template shipped in
Microsoft's office-addin-dev-settings (MIT): a web extension referencing the
registry-sideloaded add-in, plus a visible task pane that points at it.
"""

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

WEBEXT = "xl/webextensions/webextension.xml"
TASKPANES = "xl/webextensions/taskpanes.xml"
TASKPANES_RELS = "xl/webextensions/_rels/taskpanes.xml.rels"
TASKPANES_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"


def read_manifest(manifest_path: Path) -> tuple[str, str]:
    text = manifest_path.read_text(encoding="utf-8")
    addin_id = re.search(r"<Id>([^<]+)</Id>", text).group(1).strip()
    version = re.search(r"<Version>([^<]+)</Version>", text).group(1).strip()
    return addin_id, version


def _parts(addin_id: str, version: str) -> dict[str, str]:
    return {
        WEBEXT: (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" '
            f'id="{{{addin_id}}}">'
            f'<we:reference id="{addin_id}" version="{version}" store="developer" storeType="Registry"/>'
            "<we:alternateReferences/><we:properties/><we:bindings/><we:snapshot/></we:webextension>"
        ),
        TASKPANES: (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">'
            '<wetp:taskpane dockstate="right" visibility="1" width="400" row="1">'
            '<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdWebExt1"/>'
            "</wetp:taskpane></wetp:taskpanes>"
        ),
        TASKPANES_RELS: (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdWebExt1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" '
            'Target="webextension.xml"/></Relationships>'
        ),
    }


def embed_taskpane(xlsx_path: Path, addin_id: str, version: str) -> None:
    """Rewrite xlsx_path in place with the task pane parts added (idempotent)."""
    with zipfile.ZipFile(xlsx_path) as src:
        names = src.namelist()
        if TASKPANES in names:
            return
        files = {name: src.read(name) for name in names}

    content_types = files["[Content_Types].xml"].decode("utf-8")
    overrides = (
        '<Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>'
        '<Override PartName="/xl/webextensions/webextension.xml" ContentType="application/vnd.ms-office.webextension+xml"/>'
    )
    files["[Content_Types].xml"] = content_types.replace("</Types>", overrides + "</Types>").encode("utf-8")

    root_rels = files["_rels/.rels"].decode("utf-8")
    rel = f'<Relationship Id="rIdTaskpanes1" Type="{TASKPANES_REL_TYPE}" Target="xl/webextensions/taskpanes.xml"/>'
    files["_rels/.rels"] = root_rels.replace("</Relationships>", rel + "</Relationships>").encode("utf-8")

    for name, xml in _parts(addin_id, version).items():
        files[name] = xml.encode("utf-8")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=xlsx_path.parent) as tmp:
        tmp_path = Path(tmp.name)
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, data in files.items():
            dst.writestr(name, data)
    shutil.move(tmp_path, xlsx_path)
