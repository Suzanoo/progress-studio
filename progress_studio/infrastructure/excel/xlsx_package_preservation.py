from __future__ import annotations

from pathlib import Path
import os
import tempfile
import zipfile


# These parts are intentionally opaque to Progress Studio during extension-only
# workbook operations such as Payment-Breakdown.  They belong to existing
# workbook features and must survive byte-for-byte.
_OPAQUE_PREFIXES = (
    "xl/externalLinks/",
    "xl/drawings/",
)


def restore_opaque_workbook_parts(
    source_workbook: Path,
    target_workbook: Path,
    *,
    prefixes: tuple[str, ...] = _OPAQUE_PREFIXES,
) -> tuple[str, ...]:
    """Restore selected OOXML package parts from source into target byte-for-byte.

    The target is expected to be a valid workbook already written by openpyxl.
    Only parts that existed in the source and match `prefixes` are replaced.
    All other target parts, including newly-created worksheets, are kept.
    """
    source = Path(source_workbook)
    target = Path(target_workbook)

    with zipfile.ZipFile(source, "r") as source_zip:
        source_parts = {
            name: source_zip.read(name)
            for name in source_zip.namelist()
            if any(name.startswith(prefix) for prefix in prefixes)
        }

    if not source_parts:
        return ()

    fd, temp_name = tempfile.mkstemp(
        prefix=f".{target.stem}.opaque.",
        suffix=target.suffix,
        dir=target.parent,
    )
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        with zipfile.ZipFile(target, "r") as target_zip, zipfile.ZipFile(
            temp_path,
            "w",
        ) as output_zip:
            restored: list[str] = []
            existing = set(target_zip.namelist())

            for info in target_zip.infolist():
                data = source_parts.get(info.filename)
                if data is not None:
                    restored.append(info.filename)
                else:
                    data = target_zip.read(info.filename)
                output_zip.writestr(info, data)

            # Defensive: if openpyxl omitted an opaque source part entirely,
            # put it back as well.  Content-type/relationship entries for these
            # pre-existing parts remain in the normal workbook package.
            for name, data in source_parts.items():
                if name in existing:
                    continue
                output_zip.writestr(name, data)
                restored.append(name)

        os.replace(temp_path, target)
        return tuple(restored)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise
