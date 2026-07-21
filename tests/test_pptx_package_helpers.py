"""Self-tests for the package-research test helpers."""

from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile

import pytest

from pptx import Presentation

from .unitutil.pptx_package import (
    ZipMemberDiff,
    assert_changed_parts,
    assert_xml_semantically_equal,
    dangling_relationship_targets,
    diff_zip_members,
    missing_relationship_references,
    save_reopen,
    save_to_bytes,
    xml_semantically_equal,
    zip_member_map,
)

_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def test_save_to_bytes_and_reopen_are_an_independent_round_trip():
    prs = Presentation()
    prs.core_properties.title = "before save"

    saved = save_to_bytes(prs)
    reopened = save_reopen(prs)
    assert reopened.core_properties.title == "before save"

    reopened.core_properties.title = "after reopen"

    assert saved.startswith(b"PK")
    assert reopened.core_properties.title == "after reopen"
    assert prs.core_properties.title == "before save"


def test_zip_member_map_rejects_duplicate_names():
    with pytest.warns(UserWarning, match="Duplicate name"):
        package = _zip_bytes_with_duplicate_member()

    with pytest.raises(AssertionError, match="duplicate ZIP member paths") as error:
        zip_member_map(package)

    assert "ppt/presentation.xml" in str(error.value)


@pytest.mark.parametrize("member_name", ["/ppt/presentation.xml", "ppt/../bad.xml", "ppt\\bad.xml"])
def test_zip_member_map_rejects_invalid_package_member_paths(member_name):
    package = _zip_bytes({member_name: b"invalid path"})

    with pytest.raises(AssertionError, match="invalid ZIP member paths") as error:
        zip_member_map(package)

    assert repr(member_name) in str(error.value)


def test_diff_zip_members_detects_sorted_added_removed_and_changed_members():
    before = _zip_bytes({"z.xml": b"same", "removed.xml": b"old", "changed.xml": b"old"})
    after = _zip_bytes({"a.xml": b"added", "z.xml": b"same", "changed.xml": b"new"})

    assert diff_zip_members(before, after) == ZipMemberDiff(
        added=("a.xml",),
        removed=("removed.xml",),
        changed=("changed.xml",),
    )
    assert_changed_parts(
        before,
        after,
        added=("a.xml",),
        removed=("removed.xml",),
        changed=("changed.xml",),
    )


def test_exact_changed_part_budget_accepts_an_unchanged_package():
    package = _zip_bytes({"unchanged.xml": b"same"})

    assert_changed_parts(package, package)


def test_exact_changed_part_budget_reports_missing_and_unexpected_changes():
    before = _zip_bytes({"actual.xml": b"old", "missing.xml": b"same"})
    after = _zip_bytes({"actual.xml": b"new", "missing.xml": b"same"})

    with pytest.raises(AssertionError, match="changed-part budget mismatch") as error:
        assert_changed_parts(before, after, changed=("missing.xml",))

    assert "missing expected ('missing.xml',)" in str(error.value)
    assert "unexpected ('actual.xml',)" in str(error.value)


def test_changed_part_budget_includes_relationships_and_content_types():
    before = _zip_bytes({"[Content_Types].xml": b"one", "_rels/.rels": b"one"})
    after = _zip_bytes({"[Content_Types].xml": b"two", "_rels/.rels": b"two"})

    assert diff_zip_members(before, after).changed == (
        "[Content_Types].xml",
        "_rels/.rels",
    )


def test_changed_part_budget_compares_raw_xml_bytes_not_semantics():
    left = b'<root first="1" second="2"/>'
    right = b'<root second="2" first="1"/>'
    before = _zip_bytes({"part.xml": left, "unchanged.bin": b"same"})
    after = _zip_bytes({"part.xml": right, "unchanged.bin": b"same"})

    assert xml_semantically_equal("part.xml", left, right)
    assert diff_zip_members(before, after) == ZipMemberDiff((), (), ("part.xml",))


def test_package_root_relationship_target_is_resolved_from_package_root():
    members = {
        "_rels/.rels": _relationships_xml(target="ppt/presentation.xml"),
        "ppt/presentation.xml": b"<presentation/>",
    }

    assert dangling_relationship_targets(members) == ()


@pytest.mark.parametrize("target", ["../media/missing.png", "/ppt/media/missing.png"])
def test_dangling_relationship_targets_reports_relative_and_root_relative_targets(target):
    members = {
        "ppt/slides/_rels/slide1.xml.rels": _relationships_xml(target=target),
        "ppt/slides/slide1.xml": b"<slide/>",
    }

    assert "ppt/media/missing.png" in dangling_relationship_targets(members)[0]


@pytest.mark.parametrize(
    ("target_attribute", "diagnostic"),
    [("", "missing Target"), (' Target=""', "empty Target")],
)
def test_relationship_target_diagnostics_cover_missing_and_empty_targets(
    target_attribute, diagnostic
):
    rels_xml = (
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">'
        f'<Relationship Id="rId1" Type="type"{target_attribute}/>'
        "</Relationships>"
    ).encode()

    assert diagnostic in dangling_relationship_targets({"_rels/.rels": rels_xml})[0]


def test_external_relationship_targets_are_ignored():
    members = {
        "_rels/.rels": _relationships_xml(
            target="https://example.com/missing", target_mode="External"
        )
    }

    assert dangling_relationship_targets(members) == ()


def test_valid_internal_relationship_targets_are_accepted():
    members = {
        "ppt/slides/_rels/slide1.xml.rels": _relationships_xml(target="../media/image1.png"),
        "ppt/media/image1.png": b"image",
    }

    assert dangling_relationship_targets(members) == ()


def test_relationship_target_dot_segments_are_normalized_without_escaping_package():
    members = {
        "ppt/slides/_rels/slide1.xml.rels": _relationships_xml(
            target="../media/../media/image1.png"
        ),
        "ppt/media/image1.png": b"image",
    }

    assert dangling_relationship_targets(members) == ()


def test_relationship_target_cannot_escape_package_root():
    members = {"_rels/.rels": _relationships_xml(target="../outside.xml")}

    issues = dangling_relationship_targets(members)

    assert len(issues) == 1
    assert "rId1" in issues[0]
    assert "escapes the package root" in issues[0]


def test_duplicate_relationship_ids_are_reported_without_being_collapsed():
    members = {
        "ppt/slides/_rels/slide1.xml.rels": _relationships_xml(
            relationships=(("rId1", "../media/one.png"), ("rId1", "../media/two.png"))
        ),
        "ppt/media/one.png": b"one",
        "ppt/media/two.png": b"two",
    }

    issues = dangling_relationship_targets(members)

    assert len(issues) == 1
    assert "ppt/slides/_rels/slide1.xml.rels" in issues[0]
    assert "duplicate relationship Id 'rId1'" in issues[0]


def test_duplicate_relationship_ids_are_reported_by_reference_scan():
    members = {
        "ppt/slides/slide1.xml": (
            f'<p:sld xmlns:p="urn:p" xmlns:r="{_OFFICE_REL_NS}" r:id="rId1"/>'
        ).encode(),
        "ppt/slides/_rels/slide1.xml.rels": _relationships_xml(
            relationships=(("rId1", "../media/one.png"), ("rId1", "../media/two.png"))
        ),
    }

    issues = missing_relationship_references(members)

    assert issues == ("ppt/slides/_rels/slide1.xml.rels: duplicate relationship Id 'rId1'",)


@pytest.mark.parametrize(
    "scanner", [dangling_relationship_targets, missing_relationship_references]
)
def test_malformed_relationship_xml_fails_with_member_context(scanner):
    members = {
        "ppt/slides/slide1.xml": b"<slide/>",
        "ppt/slides/_rels/slide1.xml.rels": b"<Relationships",
    }

    with pytest.raises(AssertionError, match="malformed XML") as error:
        scanner(members)

    assert "ppt/slides/_rels/slide1.xml.rels" in str(error.value)


def test_wrong_relationships_namespace_fails_clearly():
    members = {"_rels/.rels": b'<Relationships xmlns="urn:not-opc"/>'}

    with pytest.raises(AssertionError, match="expected Relationships root") as error:
        dangling_relationship_targets(members)

    assert "_rels/.rels" in str(error.value)


@pytest.mark.parametrize("attribute", ["id", "embed", "link"])
def test_undefined_office_relationship_attributes_are_reported(attribute):
    members = {
        "ppt/slides/slide1.xml": (
            f'<p:sld xmlns:p="urn:p" xmlns:r="{_OFFICE_REL_NS}">'
            f'<p:node r:{attribute}="rId9"/></p:sld>'
        ).encode(),
        "ppt/slides/_rels/slide1.xml.rels": _relationships_xml(
            target="../media/image1.png", relationship_id="rId1"
        ),
    }

    issues = missing_relationship_references(members)

    assert len(issues) == 1
    assert f"r:{attribute}" in issues[0]
    assert "rId9" in issues[0]


def test_defined_office_relationship_attributes_are_accepted():
    members = {
        "ppt/slides/slide1.xml": (
            f'<p:sld xmlns:p="urn:p" xmlns:r="{_OFFICE_REL_NS}">'
            '<p:node r:id="rId1" r:embed="rId2" r:link="rId3"/></p:sld>'
        ).encode(),
        "ppt/slides/_rels/slide1.xml.rels": _relationships_xml(
            relationships=(
                ("rId1", "../slideLayouts/slideLayout1.xml"),
                ("rId2", "../media/image1.png"),
                ("rId3", "../slides/slide2.xml"),
            )
        ),
    }

    assert missing_relationship_references(members) == ()


def test_relationship_reference_without_an_owning_rels_part_is_reported():
    members = {
        "ppt/slides/slide1.xml": (
            f'<p:sld xmlns:p="urn:p" xmlns:r="{_OFFICE_REL_NS}" r:id="rId1"/>'
        ).encode()
    }

    issues = missing_relationship_references(members)

    assert issues == ("ppt/slides/slide1.xml: r:id references undefined relationship 'rId1'",)


def test_attributes_in_unrelated_namespaces_are_not_relationship_references():
    members = {"ppt/slides/slide1.xml": b'<p:sld xmlns:p="urn:p" p:id="rId9" plain="rId8"/>'}

    assert missing_relationship_references(members) == ()


def test_malformed_source_xml_fails_relationship_reference_scan_with_member_context():
    with pytest.raises(AssertionError, match="malformed XML") as error:
        missing_relationship_references({"ppt/slides/slide1.xml": b"<slide"})

    assert "ppt/slides/slide1.xml" in str(error.value)


def test_xml_semantic_equality_ignores_declarations_prefixes_attributes_and_indentation():
    left = b"""<?xml version="1.0" encoding="UTF-8"?>
        <p:root xmlns:p="urn:p" xmlns:a="urn:a" z="2" a:name="one">
          <p:child value="x" />
        </p:root>"""
    right = b"""<x:root xmlns:x="urn:p" xmlns:y="urn:a" y:name="one" z="2"><x:child
        value="x"></x:child></x:root>"""

    assert xml_semantically_equal("ppt/slides/slide1.xml", left, right)


@pytest.mark.parametrize(
    ("left", "right"),
    [
        (b"<root><a/><b/></root>", b"<root><b/><a/></root>"),
        (b'<root value="one"/>', b'<root value="two"/>'),
        (b"<root><text>word </text></root>", b"<root><text>word</text></root>"),
    ],
)
def test_xml_semantic_equality_preserves_child_order_attributes_and_text_whitespace(left, right):
    assert not xml_semantically_equal("ppt/slides/slide1.xml", left, right)


def test_xml_semantic_equality_preserves_inline_whitespace_and_tail_text():
    left = b"<root><a/> <b/>tail </root>"
    right = b"<root><a/><b/>tail</root>"

    assert not xml_semantically_equal("ppt/slides/slide1.xml", left, right)


def test_semantic_xml_assertion_identifies_member_and_difference():
    with pytest.raises(AssertionError, match="semantic XML mismatch") as error:
        assert_xml_semantically_equal("ppt/slides/slide1.xml", b"<left/>", b"<right/>")

    diagnostic = str(error.value)
    assert "ppt/slides/slide1.xml" in diagnostic
    assert "left" in diagnostic
    assert "right" in diagnostic


def test_content_types_child_order_is_semantically_irrelevant():
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    left = (
        f'<Types xmlns="{namespace}"><Default Extension="xml" ContentType="xml"/>'
        '<Override PartName="/ppt/presentation.xml" ContentType="presentation"/></Types>'
    ).encode()
    right = (
        f'<Types xmlns="{namespace}"><Override ContentType="presentation" '
        'PartName="/ppt/presentation.xml"/><Default ContentType="xml" '
        'Extension="xml"/></Types>'
    ).encode()

    assert xml_semantically_equal("[Content_Types].xml", left, right)


def test_content_types_duplicates_and_meaningful_attributes_are_retained():
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    duplicate = (
        f'<Types xmlns="{namespace}"><Default Extension="xml" ContentType="one"/>'
        '<Default Extension="xml" ContentType="two"/></Types>'
    ).encode()
    single = (
        f'<Types xmlns="{namespace}"><Default Extension="xml" ContentType="two"/></Types>'
    ).encode()

    assert not xml_semantically_equal("[Content_Types].xml", duplicate, single)


def test_relationship_semantics_do_not_collapse_duplicate_ids():
    duplicate = _relationships_xml(relationships=(("rId1", "one.xml"), ("rId1", "two.xml")))
    single = _relationships_xml(relationships=(("rId1", "one.xml"),))

    assert not xml_semantically_equal("ppt/slides/_rels/slide1.xml.rels", duplicate, single)


def test_relationship_child_order_is_semantically_irrelevant_but_ids_are_not():
    left = _relationships_xml(relationships=(("rId1", "one.xml"), ("rId2", "two.xml")))
    reordered = _relationships_xml(relationships=(("rId2", "two.xml"), ("rId1", "one.xml")))
    remapped = _relationships_xml(relationships=(("rId9", "one.xml"), ("rId2", "two.xml")))

    assert xml_semantically_equal("ppt/slides/_rels/slide1.xml.rels", left, reordered)
    assert not xml_semantically_equal("ppt/slides/_rels/slide1.xml.rels", left, remapped)


@pytest.mark.parametrize(
    ("field", "left_value", "right_value"),
    [
        ("Id", "rId1", "rId9"),
        ("Type", "type-one", "type-two"),
        ("Target", "one.xml", "two.xml"),
        ("TargetMode", "External", "Internal"),
    ],
)
def test_relationship_semantics_retain_all_relationship_fields(field, left_value, right_value):
    def rels(value):
        attributes = {
            "Id": "rId1",
            "Type": "type-one",
            "Target": "one.xml",
            "TargetMode": "External",
        }
        attributes[field] = value
        serialized = " ".join(f'{name}="{item}"' for name, item in attributes.items())
        return (
            f'<Relationships xmlns="{_PACKAGE_REL_NS}"><Relationship {serialized}/></Relationships>'
        ).encode()

    assert not xml_semantically_equal(
        "ppt/slides/_rels/slide1.xml.rels", rels(left_value), rels(right_value)
    )


def test_relationship_semantics_do_not_hide_unrelated_child_order_changes():
    left = (
        f'<Relationships xmlns="{_PACKAGE_REL_NS}" xmlns:x="urn:x">'
        "<x:first/><x:second/></Relationships>"
    ).encode()
    right = (
        f'<Relationships xmlns="{_PACKAGE_REL_NS}" xmlns:x="urn:x">'
        "<x:second/><x:first/></Relationships>"
    ).encode()

    assert not xml_semantically_equal("ppt/slides/_rels/slide1.xml.rels", left, right)


def _relationships_xml(
    *,
    target: str = "target.xml",
    target_mode: str | None = None,
    relationship_id: str = "rId1",
    relationships: tuple[tuple[str, str], ...] | None = None,
) -> bytes:
    relationship_specs = relationships or ((relationship_id, target),)
    mode_attribute = "" if target_mode is None else f' TargetMode="{target_mode}"'
    children = "".join(
        f'<Relationship Id="{rel_id}" Type="type" Target="{rel_target}"{mode_attribute}/>'
        for rel_id, rel_target in relationship_specs
    )
    return f'<Relationships xmlns="{_PACKAGE_REL_NS}">{children}</Relationships>'.encode()


def _zip_bytes(members: dict[str, bytes]) -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w") as zip_file:
        for member_name, blob in members.items():
            zip_file.writestr(member_name, blob)
    return stream.getvalue()


def _zip_bytes_with_duplicate_member() -> bytes:
    stream = BytesIO()
    with ZipFile(stream, "w") as zip_file:
        zip_file.writestr("ppt/presentation.xml", b"first")
        zip_file.writestr("ppt/presentation.xml", b"second")
    return stream.getvalue()
