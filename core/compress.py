"""Size compression for .xlsx / .xlsm workbooks.

An .xlsx file is a ZIP package. Instead of re-exporting the workbook through
openpyxl (which would silently DROP images, charts and formatting), we edit
the package in place:

* re-encode embedded images (JPEG: lower quality; PNG: re-compress),
* remove always-safe junk (calc chain, document thumbnail),
* preset "max": additionally strip conditional formatting / data validations
  and remove genuinely empty sheets (with full relationship cleanup),
* repack everything with maximum DEFLATE compression.

Memory design: only the small XML parts that may need editing (content
types, relationships, workbook.xml, sheet XMLs under a size guard) are held
in memory. Everything else — including large media files — is streamed
directly from the input archive to the output archive, so even multi-hundred-
MB workbooks compress within a few hundred MB of RAM.
"""
from __future__ import annotations

import io
import os
import posixpath
import shutil
import zipfile
import xml.etree.ElementTree as ET

from .errors import ExcelToolError

# XML namespaces used by Office Open XML
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

# Parts that must never be removed
PROTECTED = {
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
}

# Sheets whose XML is bigger than this are left untouched (memory guard)
MAX_SHEET_XML_BYTES = 12 * 1024 * 1024

# Entry types we hold in memory (everything else is streamed)
def _should_load_in_memory(name: str) -> bool:
    low = name.lower()
    if name == "[Content_Types].xml" or name == "xl/workbook.xml":
        return True
    if low.endswith(".rels"):
        return True
    if low.startswith("xl/worksheets/") and low.endswith(".xml"):
        return True
    return False


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _serialize(root: ET.Element, default_ns: str) -> bytes:
    """Serialise an ElementTree keeping the given namespace as the default."""
    ET.register_namespace("", default_ns)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _recompress_image(data: bytes, ext: str, max_dim: int, jpeg_quality: int):
    """Return smaller image bytes, or None if nothing could be saved."""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        return None

    try:
        resample = getattr(Image, "Resampling", Image).LANCZOS
        if max_dim and max(im.size) > max_dim:
            im.thumbnail((max_dim, max_dim), resample)
        buf = io.BytesIO()
        if ext in (".jpg", ".jpeg"):
            if im.mode in ("RGBA", "LA", "P"):
                rgba = im.convert("RGBA")
                bg = Image.new("RGB", rgba.size, (255, 255, 255))
                bg.paste(rgba, mask=rgba.split()[-1])
                im = bg
            elif im.mode != "RGB":
                im = im.convert("RGB")
            im.save(buf, "JPEG", quality=jpeg_quality, optimize=True, progressive=True)
        else:  # .png
            im.save(buf, "PNG", optimize=True, compress_level=9)
        out = buf.getvalue()
        return out if len(out) < len(data) else None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Relationship helpers (all operate on the in-memory "pool" of small parts)
# ---------------------------------------------------------------------------
def _rels_path_for(part: str) -> str:
    if "/" in part:
        d, b = part.rsplit("/", 1)
        return f"{d}/_rels/{b}.rels"
    return f"_rels/{part}.rels"


def _resolve_target(rel: ET.Element, rels_name: str):
    """Resolve a Relationship Target to a package-relative part name."""
    if rel.get("TargetMode") == "External":
        return None
    t = (rel.get("Target") or "").split("#", 1)[0]
    if not t:
        return None
    if t.startswith("/"):
        full = posixpath.normpath(t.lstrip("/"))
    elif "/" in rels_name:
        dirn = rels_name.rsplit("/", 1)[0]
        full = posixpath.normpath(f"{dirn}/{t}")
    else:
        full = posixpath.normpath(t)
    return full or None


def _rels_targets(rels_xml: bytes, rels_name: str):
    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return []
    out = []
    for rel in root:
        t = _resolve_target(rel, rels_name)
        if t:
            out.append(t)
    return out


def _cascade_removed(pool: dict, removed: set):
    """Add rels files (and everything they reference) of removed parts."""
    queue = list(removed)
    while queue:
        part = queue.pop(0)
        rels = _rels_path_for(part)
        if rels not in pool or rels in removed:
            continue
        removed.add(rels)
        for target in _rels_targets(pool[rels], rels):
            if target in pool and target not in removed:
                removed.add(target)
                queue.append(target)


def _collect_referenced(pool: dict, removed: set) -> set:
    """Parts still referenced by any remaining .rels file."""
    referenced = set()
    for name in pool:
        if name in removed or not name.lower().endswith(".rels"):
            continue
        referenced.update(_rels_targets(pool[name], name))
    return referenced


def _strip_dangling_rels(pool: dict, final_removed: set):
    """Remove Relationship elements pointing at parts we deleted."""
    for name in list(pool):
        if not name.lower().endswith(".rels"):
            continue
        try:
            root = ET.fromstring(pool[name])
        except ET.ParseError:
            continue
        changed = False
        for rel in list(root):
            if _resolve_target(rel, name) in final_removed:
                root.remove(rel)
                changed = True
        if changed:
            pool[name] = _serialize(root, NS_REL)


def _strip_content_type_overrides(pool: dict, final_removed: set):
    """Remove Content-Type Overrides for deleted parts."""
    name = "[Content_Types].xml"
    try:
        root = ET.fromstring(pool[name])
    except ET.ParseError:
        return
    changed = False
    for ov in list(root):
        if _localname(ov.tag) != "Override":
            continue
        pn = (ov.get("PartName") or "").lstrip("/")
        if pn in final_removed:
            root.remove(ov)
            changed = True
    if changed:
        pool[name] = _serialize(root, NS_CT)


# ---------------------------------------------------------------------------
# Preset "max": strip junk inside sheets + remove empty sheets
# ---------------------------------------------------------------------------
def _apply_max_preset(pool: dict, removed: set, stats: dict):
    rels_path = "xl/_rels/workbook.xml.rels"
    rid2target = {}
    if rels_path in pool:
        try:
            root = ET.fromstring(pool[rels_path])
        except ET.ParseError:
            root = None
        if root is not None:
            for rel in root:
                rid = rel.get("Id")
                t = _resolve_target(rel, rels_path)
                if rid and t:
                    rid2target[rid] = t

    try:
        wbroot = ET.fromstring(pool["xl/workbook.xml"])
    except ET.ParseError:
        wbroot = None

    sheet_infos = []  # (name, rId, part_path)
    if wbroot is not None:
        sheets_el = wbroot.find(f"{{{NS_MAIN}}}sheets")
        if sheets_el is not None:
            for sh in sheets_el:
                name = sh.get("name") or "?"
                rid = sh.get(f"{{{NS_REL}}}id")
                part = rid2target.get(rid or "")
                if part and part in pool:
                    sheet_infos.append((name, rid, part))

    remove_sheets = []
    stripped = 0
    for name, rid, part in sheet_infos:
        if part not in pool:
            continue
        xml = pool[part]
        if len(xml) > MAX_SHEET_XML_BYTES:
            continue  # leave giant sheets untouched (memory guard)
        try:
            root = ET.fromstring(xml)
        except ET.ParseError:
            continue

        # -- empty sheet detection -----------------------------------------
        # A sheet is "empty" when no row contains a real value, formula or
        # shared-string reference — phantom rows with blank cells don't count.
        def _row_has_content(row_el):
            for cell in row_el:
                if _localname(cell.tag) != "c":
                    continue
                for child in cell:
                    if _localname(child.tag) in ("v", "is", "f"):
                        return True
            return False

        sd = root.find(f".//{{{NS_MAIN}}}sheetData")
        has_rows = sd is not None and any(
            _localname(el.tag) == "row" and _row_has_content(el) for el in sd
        )
        has_artifacts = any(
            _localname(el.tag) in {"drawing", "legacyDrawing", "oleObjects", "controls"}
            for el in root.iter()
        )
        if not has_rows and not has_artifacts:
            remove_sheets.append((name, rid, part))
            continue

        # -- strip conditional formatting / data validations ---------------
        before = len(xml)
        for el in list(root.iter()):
            if _localname(el.tag) in {"conditionalFormatting", "dataValidations"}:
                root.remove(el)
        after = len(_serialize(root, NS_MAIN))
        if after < before:
            pool[part] = _serialize(root, NS_MAIN)
            stripped += 1

    # never leave a workbook with zero sheets
    if remove_sheets and len(remove_sheets) >= len(sheet_infos):
        remove_sheets = []

    for sheet_name, rid, part in remove_sheets:
        removed.add(part)
        stats["empty_sheets_removed"].append(sheet_name)

    if remove_sheets and wbroot is not None:
        sheets_el = wbroot.find(f"{{{NS_MAIN}}}sheets")
        if sheets_el is not None:
            removed_rids = {rid for _, rid, _ in remove_sheets}
            for sh in list(sheets_el):
                if sh.get(f"{{{NS_REL}}}id") in removed_rids:
                    sheets_el.remove(sh)
            pool["xl/workbook.xml"] = _serialize(wbroot, NS_MAIN)

    stats["sheets_stripped"] = stripped


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def compress_workbook(src_path: str, dst_path: str, *,
                      preset: str = "safe", max_dim: int = 1600,
                      jpeg_quality: int = 72) -> dict:
    """Compress an .xlsx/.xlsm workbook and write the result to dst_path.

    Returns a stats dict. Never raises on cosmetic problems; raises
    ExcelToolError only for genuinely invalid input.
    """
    if preset not in ("safe", "max"):
        preset = "safe"
    max_dim = max(0, min(int(max_dim), 4096))
    jpeg_quality = max(30, min(int(jpeg_quality), 95))

    try:
        zin = zipfile.ZipFile(src_path, "r")
    except zipfile.BadZipFile as exc:
        raise ExcelToolError(
            "Could not read the workbook: the file is not a valid .xlsx/.xlsm "
            "package (it may be corrupted or password-protected)."
        ) from exc

    try:
        entries = [e for e in zin.infolist() if not e.is_dir()]
        names = [e.filename for e in entries]
        if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
            raise ExcelToolError(
                "Could not read the workbook: required package parts are missing "
                "(is this really an .xlsx/.xlsm file?)."
            )

        stats = {
            "images": 0,
            "images_compressed": 0,
            "image_bytes_saved": 0,
            "empty_sheets_removed": [],
            "sheets_stripped": 0,
            "removed_parts": [],
        }
        removed = set()
        pool = {}  # small XML parts that may be edited, held in memory

        # ---- 1. load the small editable parts ------------------------------
        for e in entries:
            name = e.filename
            low = name.lower()
            if low == "xl/calcchain.xml" or low.startswith("docprops/thumbnail"):
                removed.add(name)  # always-safe junk; never loaded
            elif _should_load_in_memory(name):
                if e.file_size <= MAX_SHEET_XML_BYTES:
                    pool[name] = zin.read(name)

        # ---- 2. preset-specific work -----------------------------------------
        if preset == "max":
            _apply_max_preset(pool, removed, stats)

        # ---- 3. cascade removals & protection ---------------------------------
        _cascade_removed(pool, removed)
        removed -= PROTECTED

        # Drop relationships that point at removed parts (e.g. the workbook
        # rels entry for a deleted sheet) BEFORE computing what is still
        # referenced — otherwise a deleted part looks "referenced" by its
        # own stale relationship and stays in the package as dead weight.
        _strip_dangling_rels(pool, removed)
        referenced = _collect_referenced(pool, removed)
        final_removed = removed - referenced

        _strip_content_type_overrides(pool, final_removed)

        for name in final_removed:
            pool.pop(name, None)
        stats["removed_parts"] = sorted(final_removed)

        # ---- 4. repack with maximum compression (streaming) -------------------
        src_size = os.path.getsize(src_path)
        with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED,
                             compresslevel=9) as zout:
            for e in entries:
                name = e.filename
                if name in final_removed:
                    continue
                if name in pool:
                    zout.writestr(name, pool[name])
                    continue
                low = name.lower()
                if low.startswith("xl/media/"):
                    ext = os.path.splitext(name)[1].lower()
                    if ext in (".png", ".jpg", ".jpeg"):
                        data = zin.read(name)
                        stats["images"] += 1
                        new = _recompress_image(data, ext, max_dim, jpeg_quality)
                        if new:
                            stats["images_compressed"] += 1
                            stats["image_bytes_saved"] += len(data) - len(new)
                            zout.writestr(name, new)
                            continue
                # stream everything else (no full-file buffering)
                with zin.open(name) as srcf, zout.open(name, "w") as dstf:
                    shutil.copyfileobj(srcf, dstf, length=1 << 20)

        dst_size = os.path.getsize(dst_path)
        if dst_size >= src_size:
            # repacking gained nothing — hand back the original byte-for-byte
            shutil.copyfile(src_path, dst_path)
            dst_size = src_size

        stats["original_bytes"] = src_size
        stats["final_bytes"] = dst_size
        stats["saved_bytes"] = src_size - dst_size
        stats["percent_saved"] = round(100.0 * (src_size - dst_size) / src_size, 1) if src_size else 0.0
        return stats
    finally:
        zin.close()
