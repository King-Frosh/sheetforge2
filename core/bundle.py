"""Bundle several files into a single .zip archive."""
from __future__ import annotations

import os
import zipfile


def make_bundle(items, output_path: str) -> dict:
    """Zip `items` (list of (source_path, display_name)) into output_path.

    Duplicate archive names are disambiguated with " (2)", " (3)" …
    """
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        used = set()
        for src, display in items:
            arc = display or os.path.basename(src)
            base, ext = os.path.splitext(arc)
            n = 2
            while arc.lower() in used:
                arc = f"{base} ({n}){ext}"
                n += 1
            used.add(arc.lower())
            z.write(src, arc)

    return {
        "files": len(items),
        "total_bytes": sum(os.path.getsize(s) for s, _ in items),
        "zip_bytes": os.path.getsize(output_path),
    }
