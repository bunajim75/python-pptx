"""Test helpers for inspecting serialized PowerPoint packages.

Package-diff and raw-relationship scanning concepts are adapted from the
MIT-licensed ``paper-instruments/paper-pptx`` test harness. This version is
independently hardened for the invariants exercised by this repository.
"""

from __future__ import annotations

import posixpath
from collections import Counter
from io import BytesIO
from typing import Iterable, NamedTuple
from urllib.parse import urlsplit
from xml.etree import ElementTree
from zipfile import ZipFile, ZipInfo

from pptx import Presentation

_CONTENT_TYPES_MEMBER = "[Content_Types].xml"
_OFFICE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class ZipMemberDiff(NamedTuple):
    """Byte-level differences between two ZIP packages."""

    added: tuple[str, ...]
    removed: tuple[str, ...]
    changed: tuple[str, ...]


class PackageRelationship(NamedTuple):
    """Raw fields retained from one OPC ``Relationship`` element."""

    relationship_id: str | None
    relationship_type: str | None
    target: str | None
    target_mode: str | None


def save_to_bytes(prs) -> bytes:
    """Serialize `prs` to an in-memory PPTX package."""
    stream = BytesIO()
    prs.save(stream)
    return stream.getvalue()


def save_reopen(prs):
    """Serialize and reopen `prs` from the resulting bytes."""
    return Presentation(BytesIO(save_to_bytes(prs)))


def zip_member_map(pptx_bytes: bytes) -> dict[str, bytes]:
    """Return package members keyed by ZIP path, rejecting duplicate paths."""
    with ZipFile(BytesIO(pptx_bytes)) as zip_file:
        members = zip_file.infolist()
        name_counts = Counter(info.filename for info in members)
        duplicate_names = tuple(sorted(name for name, count in name_counts.items() if count > 1))
        if duplicate_names:
            raise AssertionError(f"duplicate ZIP member paths: {duplicate_names!r}")
        invalid_names = tuple(
            sorted(info.filename for info in members if not _is_valid_package_member(info))
        )
        if invalid_names:
            raise AssertionError(f"invalid ZIP member paths: {invalid_names!r}")
        return {info.filename: zip_file.read(info) for info in members}


def diff_zip_members(before_bytes: bytes, after_bytes: bytes) -> ZipMemberDiff:
    """Return deterministic byte-level package-member differences."""
    before = zip_member_map(before_bytes)
    after = zip_member_map(after_bytes)
    before_names = set(before)
    after_names = set(after)
    common_names = before_names & after_names
    return ZipMemberDiff(
        added=tuple(sorted(after_names - before_names)),
        removed=tuple(sorted(before_names - after_names)),
        changed=tuple(sorted(name for name in common_names if before[name] != after[name])),
    )


def assert_changed_parts(
    before_bytes: bytes,
    after_bytes: bytes,
    *,
    added: Iterable[str] = (),
    removed: Iterable[str] = (),
    changed: Iterable[str] = (),
) -> None:
    """Assert an exact byte-level budget for package-member changes."""
    expected = ZipMemberDiff(
        tuple(sorted(set(added))),
        tuple(sorted(set(removed))),
        tuple(sorted(set(changed))),
    )
    actual = diff_zip_members(before_bytes, after_bytes)
    if actual == expected:
        return

    details: list[str] = []
    for category in ZipMemberDiff._fields:
        expected_names = set(getattr(expected, category))
        actual_names = set(getattr(actual, category))
        missing = tuple(sorted(expected_names - actual_names))
        unexpected = tuple(sorted(actual_names - expected_names))
        if missing:
            details.append(f"{category}: missing expected {missing!r}")
        if unexpected:
            details.append(f"{category}: unexpected {unexpected!r}")
    raise AssertionError("changed-part budget mismatch:\n" + "\n".join(details))


def dangling_relationship_targets(member_map: dict[str, bytes]) -> tuple[str, ...]:
    """Return diagnostics for invalid or dangling internal relationship targets."""
    issues: list[str] = []
    for rels_member in sorted(name for name in member_map if name.endswith(".rels")):
        relationships = _relationships(rels_member, member_map[rels_member])
        issues.extend(_duplicate_relationship_id_issues(rels_member, relationships))
        for relationship in relationships:
            relationship_id = relationship.relationship_id or "<missing Id>"
            if relationship.relationship_id is None or not relationship.relationship_id.strip():
                issues.append(f"{rels_member}: relationship has missing or empty Id")
            if relationship.relationship_type is None or not relationship.relationship_type.strip():
                issues.append(
                    f"{rels_member}: relationship {relationship_id} has missing or empty Type"
                )
            if relationship.target_mode == "External":
                continue
            if relationship.target_mode not in (None, "Internal"):
                issues.append(
                    f"{rels_member}: relationship {relationship_id} has invalid TargetMode "
                    f"{relationship.target_mode!r}"
                )
                continue
            target = relationship.target
            if target is None:
                issues.append(f"{rels_member}: relationship {relationship_id} has missing Target")
                continue
            if not target.strip():
                issues.append(f"{rels_member}: relationship {relationship_id} has empty Target")
                continue
            try:
                resolved_target = _resolve_relationship_target(rels_member, target)
            except ValueError as error:
                issues.append(f"{rels_member}: relationship {relationship_id}: {error}")
                continue
            if resolved_target not in member_map:
                issues.append(
                    f"{rels_member}: relationship {relationship_id} targets missing part "
                    f"{resolved_target!r} (Target={target!r})"
                )
    return tuple(sorted(issues))


def missing_relationship_references(member_map: dict[str, bytes]) -> tuple[str, ...]:
    """Return Office relationship attributes whose IDs are not defined."""
    issues: list[str] = []
    relationship_attribute_prefix = f"{{{_OFFICE_RELATIONSHIPS_NS}}}"
    for member_name in sorted(member_map):
        if not _is_xml_member(member_name) or member_name.endswith(".rels"):
            continue
        root = _parse_xml(member_name, member_map[member_name])
        defined_ids, relationship_issues = _defined_relationship_ids(member_name, member_map)
        issues.extend(relationship_issues)
        for element in root.iter():
            for attribute_name, relationship_id in sorted(element.attrib.items()):
                if not attribute_name.startswith(relationship_attribute_prefix):
                    continue
                attribute_local_name = attribute_name[len(relationship_attribute_prefix) :]
                if relationship_id not in defined_ids:
                    issues.append(
                        f"{member_name}: r:{attribute_local_name} references undefined "
                        f"relationship {relationship_id!r}"
                    )
    return tuple(sorted(issues))


def xml_semantically_equal(member_name: str, left: bytes, right: bytes) -> bool:
    """Compare XML structure while ignoring serialization-only differences."""
    left_root = _parse_xml(f"{member_name} (left)", left)
    right_root = _parse_xml(f"{member_name} (right)", right)
    return _semantic_element(left_root, member_name) == _semantic_element(right_root, member_name)


def assert_xml_semantically_equal(member_name: str, left: bytes, right: bytes) -> None:
    """Assert semantic XML equality with member-specific mismatch diagnostics."""
    left_root = _parse_xml(f"{member_name} (left)", left)
    right_root = _parse_xml(f"{member_name} (right)", right)
    left_semantic = _semantic_element(left_root, member_name)
    right_semantic = _semantic_element(right_root, member_name)
    if left_semantic == right_semantic:
        return
    raise AssertionError(
        f"{member_name}: semantic XML mismatch\nleft:  {left_semantic!r}\nright: {right_semantic!r}"
    )


def _defined_relationship_ids(
    source_member: str, member_map: dict[str, bytes]
) -> tuple[frozenset[str], tuple[str, ...]]:
    rels_member = _rels_member_for(source_member)
    if rels_member not in member_map:
        return frozenset(), ()
    relationships = _relationships(rels_member, member_map[rels_member])
    return (
        frozenset(
            relationship.relationship_id
            for relationship in relationships
            if relationship.relationship_id is not None
        ),
        _duplicate_relationship_id_issues(rels_member, relationships),
    )


def _is_xml_member(member_name: str) -> bool:
    return member_name == _CONTENT_TYPES_MEMBER or member_name.endswith((".xml", ".rels"))


def _relationship_source_member(rels_member: str) -> str | None:
    if rels_member == "_rels/.rels":
        return None
    directory, filename = posixpath.split(rels_member)
    if posixpath.basename(directory) != "_rels" or not filename.endswith(".rels"):
        raise ValueError(f"not an OPC relationships member: {rels_member!r}")
    source_directory = posixpath.dirname(directory)
    source_filename = filename[: -len(".rels")]
    return posixpath.join(source_directory, source_filename)


def _rels_member_for(source_member: str) -> str:
    directory, filename = posixpath.split(source_member)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _resolve_relationship_target(rels_member: str, target: str) -> str:
    source_member = _relationship_source_member(rels_member)
    target_uri = urlsplit(target)
    if target_uri.scheme or target_uri.netloc:
        raise ValueError(f"internal Target is not a package path: {target!r}")
    target_path = target_uri.path
    if target_path.startswith("/"):
        candidate = target_path.lstrip("/")
    else:
        source_directory = "" if source_member is None else posixpath.dirname(source_member)
        candidate = posixpath.join(source_directory, target_path)
    resolved = posixpath.normpath(candidate)
    if resolved in ("", ".", "..") or resolved.startswith("../"):
        raise ValueError(f"internal Target escapes the package root: {target!r}")
    return resolved


def _semantic_element(element: ElementTree.Element, member_name: str):
    children = list(element)
    semantic_children = [
        (_semantic_element(child, member_name), _semantic_tail(child.tail)) for child in children
    ]
    if _children_are_order_insensitive(member_name, element.tag, children):
        semantic_children.sort(key=repr)
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        _semantic_whitespace(element.text),
        tuple(semantic_children),
    )


def _semantic_tail(tail: str | None) -> str | None:
    return _semantic_whitespace(tail)


def _children_are_order_insensitive(
    member_name: str, tag: str, children: list[ElementTree.Element] | None = None
) -> bool:
    children = [] if children is None else children
    if member_name == _CONTENT_TYPES_MEMBER:
        namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
        return tag == f"{{{namespace}}}Types" and all(
            child.tag in {f"{{{namespace}}}Default", f"{{{namespace}}}Override"}
            for child in children
        )
    if member_name.endswith(".rels"):
        return tag == f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationships" and all(
            child.tag == f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship" for child in children
        )
    return False


def _duplicate_relationship_id_issues(
    rels_member: str, relationships: tuple[PackageRelationship, ...]
) -> tuple[str, ...]:
    counts = Counter(
        relationship.relationship_id
        for relationship in relationships
        if relationship.relationship_id is not None
    )
    return tuple(
        f"{rels_member}: duplicate relationship Id {relationship_id!r}"
        for relationship_id, count in sorted(counts.items())
        if count > 1
    )


def _is_valid_package_member(info: ZipInfo) -> bool:
    name = info.filename
    return not (
        not name
        or info.is_dir()
        or name.startswith("/")
        or "\\" in name
        or any(segment in {"", ".", ".."} for segment in name.split("/"))
    )


def _parse_xml(member_name: str, blob: bytes) -> ElementTree.Element:
    try:
        return ElementTree.fromstring(blob)
    except ElementTree.ParseError as error:
        raise AssertionError(f"{member_name}: malformed XML: {error}") from error


def _relationships(rels_member: str, blob: bytes) -> tuple[PackageRelationship, ...]:
    root = _parse_xml(rels_member, blob)
    expected_root = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationships"
    expected_child = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
    if root.tag != expected_root:
        raise AssertionError(
            f"{rels_member}: expected Relationships root in package relationships namespace"
        )
    unexpected_children = tuple(child.tag for child in root if child.tag != expected_child)
    if unexpected_children:
        raise AssertionError(
            f"{rels_member}: unexpected relationship child elements {unexpected_children!r}"
        )
    return tuple(
        PackageRelationship(
            relationship.get("Id"),
            relationship.get("Type"),
            relationship.get("Target"),
            relationship.get("TargetMode"),
        )
        for relationship in root
    )


def _semantic_whitespace(value: str | None) -> str | None:
    if value is not None and value.isspace() and any(char in value for char in "\r\n\t"):
        return None
    return value
