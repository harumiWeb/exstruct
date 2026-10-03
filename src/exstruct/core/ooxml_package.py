"""Reusable OOXML package relationship helpers."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import PurePosixPath
from zipfile import ZipFile

from defusedxml import ElementTree

_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS = {"rel": _RELATIONSHIPS_NS}


@dataclass(frozen=True)
class OoxmlRelationship:
    """Relationship metadata extracted from an OOXML ``.rels`` part."""

    target: str
    relationship_type: str
    external: bool = False


def read_relationships(
    archive: ZipFile, source_path: str
) -> dict[str, OoxmlRelationship]:
    """Read relationships owned by ``source_path`` from an open OOXML archive."""

    rels_path = relationship_part_path(source_path)
    if rels_path not in archive.NameToInfo:
        return {}
    root = ElementTree.fromstring(archive.read(rels_path))
    base_dir = _base_dir(source_path)
    rel_map: dict[str, OoxmlRelationship] = {}
    for rel in root.findall("rel:Relationship", _NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        relationship_type = rel.attrib.get("Type")
        if not rel_id or not target or not relationship_type:
            continue
        external = rel.attrib.get("TargetMode") == "External"
        rel_map[rel_id] = OoxmlRelationship(
            target=target if external else normalize_part_path(base_dir, target),
            relationship_type=relationship_type,
            external=external,
        )
    return rel_map


def relationship_part_path(source_path: str) -> str:
    """Return the package path of the relationships part for a source part."""

    path = PurePosixPath(source_path)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def normalize_part_path(base_dir: str, target: str) -> str:
    """Resolve an internal POSIX target against its source part directory."""

    base = PurePosixPath(base_dir)
    normalized = base.joinpath(PurePosixPath(target)).as_posix()
    parts: list[str] = []
    for part in normalized.split("/"):
        if part in {"", "."}:
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def _source_path_from_rels(rels_path: str) -> str:
    """Recover the source part path that owns a relationships part."""

    rels = PurePosixPath(rels_path)
    if rels.parent.name != "_rels":
        return rels_path
    stem = rels.name.removesuffix(".rels")
    return str(rels.parent.parent / stem)


def _base_dir(path: str) -> str:
    """Return the POSIX parent directory for a package path."""

    return str(PurePosixPath(path).parent)
