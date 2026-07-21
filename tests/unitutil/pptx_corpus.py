"""Validation helpers for the frozen PPTX research corpus."""

from __future__ import annotations

import hashlib
import json
import re
from datetime import date
from pathlib import Path, PurePosixPath

_CONTROL_FILES = frozenset({"MANIFEST.sha256", "README.md"})
_DESKTOP_SUFFIX = ".desktop.json"
_EMPTY_CORPUS_MARKER = "# EMPTY: no fixture artifacts are committed yet."
_ROLES = frozenset({"powerpoint-authored", "powerpoint-normalized", "synthetic-invalid"})
_SCENARIO_RE = re.compile(r"[a-z0-9]+(?:-[a-z0-9]+)*")
_SHA256_RE = re.compile(r"[0-9a-f]{64}")


def sha256_file(path: Path) -> str:
    """Return the lowercase SHA-256 digest for `path`."""
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_manifest(path: Path) -> dict[str, str]:
    """Parse a SHA-256 manifest, rejecting malformed and duplicate paths."""
    entries: dict[str, str] = {}
    for line_number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not raw_line or raw_line.startswith("#"):
            continue
        match = re.fullmatch(r"([0-9a-f]{64})  (.+)", raw_line)
        if match is None:
            raise AssertionError(f"{path}:{line_number}: invalid SHA-256 manifest entry")
        digest, member_name = match.groups()
        _validate_manifest_path(path, line_number, member_name)
        if member_name in entries:
            raise AssertionError(f"{path}:{line_number}: duplicate manifest path {member_name!r}")
        entries[member_name] = digest
    return entries


def validate_fixture_corpus(corpus_root: Path) -> None:
    """Validate manifest agreement, fixture pairing, hashes, and JSON metadata."""
    manifest_path = corpus_root / "MANIFEST.sha256"
    readme_path = corpus_root / "README.md"
    if not readme_path.is_file():
        raise AssertionError(f"fixture corpus is missing control file {readme_path}")
    if not manifest_path.is_file():
        raise AssertionError(f"fixture corpus is missing control file {manifest_path}")
    entries = parse_manifest(manifest_path)
    manifest_lines = manifest_path.read_text(encoding="utf-8").splitlines()
    if not entries and _EMPTY_CORPUS_MARKER not in manifest_lines:
        raise AssertionError(f"{manifest_path}: empty corpus must contain {_EMPTY_CORPUS_MARKER!r}")
    if entries and _EMPTY_CORPUS_MARKER in manifest_lines:
        raise AssertionError(
            f"{manifest_path}: empty-corpus marker cannot accompany fixture entries"
        )
    disk_paths = {
        path.relative_to(corpus_root).as_posix()
        for path in corpus_root.rglob("*")
        if path.is_file() and path.relative_to(corpus_root).as_posix() not in _CONTROL_FILES
    }
    manifest_paths = set(entries)
    missing_files = tuple(sorted(manifest_paths - disk_paths))
    unlisted_files = tuple(sorted(disk_paths - manifest_paths))
    if missing_files or unlisted_files:
        details = []
        if missing_files:
            details.append(f"missing files: {missing_files!r}")
        if unlisted_files:
            details.append(f"unlisted files: {unlisted_files!r}")
        raise AssertionError("manifest does not exactly match corpus: " + "; ".join(details))

    for member_name, expected_digest in sorted(entries.items()):
        actual_digest = sha256_file(corpus_root / member_name)
        if actual_digest != expected_digest:
            raise AssertionError(
                f"hash mismatch for {member_name!r}: expected {expected_digest}, "
                f"got {actual_digest}"
            )

    fixture_paths = {path for path in disk_paths if path.endswith(".pptx")}
    sidecar_paths = {
        path for path in disk_paths if path.endswith(".json") and not path.endswith(_DESKTOP_SUFFIX)
    }
    desktop_paths = {path for path in disk_paths if path.endswith(_DESKTOP_SUFFIX)}
    supported_paths = fixture_paths | sidecar_paths | desktop_paths
    unsupported_paths = tuple(sorted(disk_paths - supported_paths))
    if unsupported_paths:
        raise AssertionError(f"unsupported corpus artifacts: {unsupported_paths!r}")

    for artifact_path in sorted(supported_paths):
        _validate_artifact_path(artifact_path)
    _validate_pairing(fixture_paths, sidecar_paths, desktop_paths)
    for fixture_path in sorted(fixture_paths):
        role = _role_for(fixture_path)
        stem = fixture_path[: -len(".pptx")]
        validate_fixture_sidecar(corpus_root / f"{stem}.json", expected_role=role)
        validate_desktop_verification(corpus_root / f"{stem}{_DESKTOP_SUFFIX}")


def validate_fixture_sidecar(path: Path, *, expected_role: str | None = None) -> None:
    """Validate one PPTX provenance sidecar."""
    data = _load_json_object(path)
    _require_exact_fields(
        path,
        data,
        {"schema_version", "role", "date", "description", "source", "application"},
    )
    _require_schema_version(path, data)
    role = data.get("role")
    if role not in _ROLES:
        raise AssertionError(f"{path}: invalid fixture role {role!r}")
    if expected_role is not None and role != expected_role:
        raise AssertionError(
            f"{path}: fixture role {role!r} does not match bucket {expected_role!r}"
        )
    _require_iso_date(path, "date", data.get("date"))
    for field_name in ("description", "source"):
        _require_nonempty_string(path, field_name, data.get(field_name))
    application = data.get("application")
    if role == "synthetic-invalid":
        if application is not None:
            _require_nonempty_string(path, "application", application)
    else:
        _require_nonempty_string(path, "application", application)
        if "powerpoint" not in application.lower():
            raise AssertionError(f"{path}: PowerPoint provenance must name PowerPoint")


def validate_desktop_verification(path: Path) -> None:
    """Validate one desktop-PowerPoint verification record."""
    data = _load_json_object(path)
    _require_exact_fields(
        path,
        data,
        {
            "schema_version",
            "status",
            "date",
            "application",
            "version",
            "platform",
            "notes",
        },
    )
    _require_schema_version(path, data)
    status = data.get("status")
    if status not in {"passed", "failed", "not-run"}:
        raise AssertionError(f"{path}: invalid desktop-verification status {status!r}")
    notes = data.get("notes")
    if not isinstance(notes, str):
        raise AssertionError(f"{path}: notes must be a string")
    if status == "not-run":
        for field_name in ("date", "application", "version", "platform"):
            if data.get(field_name) is not None:
                raise AssertionError(f"{path}: {field_name} must be null when status is not-run")
        return
    _require_iso_date(path, "date", data.get("date"))
    for field_name in ("application", "version", "platform"):
        _require_nonempty_string(path, field_name, data.get(field_name))
    if "powerpoint" not in data["application"].lower():
        raise AssertionError(f"{path}: desktop verification must name PowerPoint")


def _load_json_object(path: Path) -> dict[str, object]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as error:
        raise AssertionError(f"{path}: invalid JSON: {error}") from error
    if not isinstance(data, dict):
        raise AssertionError(f"{path}: JSON root must be an object")
    return data


def _require_iso_date(path: Path, field_name: str, value: object) -> None:
    if not isinstance(value, str) or re.fullmatch(r"\d{4}-\d{2}-\d{2}", value) is None:
        raise AssertionError(f"{path}: {field_name} must be an ISO date (YYYY-MM-DD)")
    try:
        date.fromisoformat(value)
    except ValueError as error:
        raise AssertionError(f"{path}: {field_name} must be an ISO date (YYYY-MM-DD)") from error


def _require_nonempty_string(path: Path, field_name: str, value: object) -> None:
    if not isinstance(value, str) or not value.strip():
        raise AssertionError(f"{path}: {field_name} must be a non-empty string")


def _require_schema_version(path: Path, data: dict[str, object]) -> None:
    if data.get("schema_version") != 1:
        raise AssertionError(f"{path}: schema_version must be 1")


def _require_exact_fields(path: Path, data: dict[str, object], required: set[str]) -> None:
    missing = tuple(sorted(required - set(data)))
    unexpected = tuple(sorted(set(data) - required))
    if missing:
        raise AssertionError(f"{path}: missing required fields {missing!r}")
    if unexpected:
        raise AssertionError(f"{path}: unexpected fields {unexpected!r}")


def _role_for(fixture_path: str) -> str:
    parts = PurePosixPath(fixture_path).parts
    role = parts[0] if len(parts) == 2 else None
    if role not in _ROLES:
        raise AssertionError(
            f"{fixture_path!r} must be stored under one of {tuple(sorted(_ROLES))!r}"
        )
    return role


def _validate_artifact_path(artifact_path: str) -> None:
    parts = PurePosixPath(artifact_path).parts
    if len(parts) != 2 or parts[0] not in _ROLES:
        raise AssertionError(
            f"{artifact_path!r} must be directly under one of {tuple(sorted(_ROLES))!r}"
        )
    filename = parts[1]
    if filename.endswith(_DESKTOP_SUFFIX):
        scenario = filename[: -len(_DESKTOP_SUFFIX)]
    else:
        scenario = Path(filename).stem
    if _SCENARIO_RE.fullmatch(scenario) is None:
        raise AssertionError(f"{artifact_path!r} scenario name must use lowercase kebab-case")


def _validate_manifest_path(path: Path, line_number: int, member_name: str) -> None:
    member_path = PurePosixPath(member_name)
    if (
        not member_name
        or member_path.is_absolute()
        or "\\" in member_name
        or any(part in {"", ".", ".."} for part in member_name.split("/"))
    ):
        raise AssertionError(f"{path}:{line_number}: invalid manifest path {member_name!r}")


def _validate_pairing(
    fixture_paths: set[str], sidecar_paths: set[str], desktop_paths: set[str]
) -> None:
    expected_sidecars = {f"{path[: -len('.pptx')]}.json" for path in fixture_paths}
    expected_desktop = {f"{path[: -len('.pptx')]}{_DESKTOP_SUFFIX}" for path in fixture_paths}
    missing_sidecars = tuple(sorted(expected_sidecars - sidecar_paths))
    orphan_sidecars = tuple(sorted(sidecar_paths - expected_sidecars))
    missing_desktop = tuple(sorted(expected_desktop - desktop_paths))
    orphan_desktop = tuple(sorted(desktop_paths - expected_desktop))
    if not (missing_sidecars or orphan_sidecars or missing_desktop or orphan_desktop):
        return
    details = []
    for label, paths in (
        ("missing fixture sidecars", missing_sidecars),
        ("orphan fixture sidecars", orphan_sidecars),
        ("missing desktop verification", missing_desktop),
        ("orphan desktop verification", orphan_desktop),
    ):
        if paths:
            details.append(f"{label}: {paths!r}")
    raise AssertionError("fixture pairing failure: " + "; ".join(details))
