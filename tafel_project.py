"""Safe, versioned on-disk project documents for the Tafel analyser.

The format deliberately contains only JSON primitives.  It is separate from the
Qt user interface so an invalid project can be rejected before any live window
state is changed.  Scientific arrays are encoded by :mod:`tafel_analyzer` as
ordinary JSON number lists; this module owns envelope validation and durable,
atomic writes.
"""
from __future__ import annotations

import json
import math
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping


PROJECT_FORMAT = "tafel-analyzer-project"
PROJECT_VERSION = 1
PROJECT_SUFFIX = ".tafelproj"
_MAX_DOCUMENT_NODES = 5_000_000
_MAX_DEPTH = 64


class ProjectError(ValueError):
    """A project-format or project-file error that can be shown to the user."""


def normalise_project_path(path: str | Path) -> Path:
    """Return a project path, adding the dedicated suffix when omitted.

    ``.tafelproj`` is intentionally distinct from CHI source files and Excel
    exports, so choosing a source filename cannot replace it accidentally.
    """
    candidate = Path(path)
    return candidate if candidate.suffix.lower() == PROJECT_SUFFIX else Path(f"{candidate}{PROJECT_SUFFIX}")


def _validate_json_value(value: Any, *, depth: int = 0, count: list[int] | None = None) -> None:
    """Reject non-JSON, non-finite, or deliberately pathological documents."""
    if depth > _MAX_DEPTH:
        raise ProjectError("项目文件层级过深，已拒绝读取。")
    if count is None:
        count = [0]
    count[0] += 1
    if count[0] > _MAX_DOCUMENT_NODES:
        raise ProjectError("项目文件过大，已拒绝读取。")
    if value is None or isinstance(value, (str, bool, int)):
        return
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ProjectError("项目文件包含非有限数值。")
        return
    if isinstance(value, list):
        for child in value:
            _validate_json_value(child, depth=depth + 1, count=count)
        return
    if isinstance(value, dict):
        for key, child in value.items():
            if not isinstance(key, str):
                raise ProjectError("项目文件包含非字符串字段名。")
            _validate_json_value(child, depth=depth + 1, count=count)
        return
    raise ProjectError("项目文件包含不支持的数据类型。")


def validate_project_document(document: object) -> dict[str, Any]:
    """Validate the stable envelope and return the JSON object unchanged.

    Detailed validation of parsed arrays and fit parameters belongs to the
    analyser because it knows the numerical invariants.  This split lets the
    UI validate the complete project into temporary model objects before it
    replaces an active session.
    """
    if not isinstance(document, dict):
        raise ProjectError("项目根节点必须是 JSON 对象。")
    _validate_json_value(document)
    if document.get("format") != PROJECT_FORMAT:
        raise ProjectError("这不是 Tafel 分析器项目文件。")
    if document.get("version") != PROJECT_VERSION:
        raise ProjectError(f"不支持的项目文件版本：{document.get('version')!r}。")
    active_view = document.get("active_view")
    if not isinstance(active_view, str) or active_view not in {"single", "batch"}:
        raise ProjectError("项目文件缺少有效的工作区标识。")
    for key in ("single", "batch"):
        value = document.get(key)
        if value is not None and not isinstance(value, dict):
            raise ProjectError(f"项目字段“{key}”格式无效。")
    return document


def make_project_document(
    *, single: Mapping[str, Any] | None, batch: Mapping[str, Any] | None, active_view: str,
) -> dict[str, Any]:
    """Construct the common project envelope from already JSON-safe payloads."""
    document: dict[str, Any] = {
        "format": PROJECT_FORMAT,
        "version": PROJECT_VERSION,
        "saved_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "active_view": active_view,
        "single": dict(single) if single is not None else None,
        "batch": dict(batch) if batch is not None else None,
    }
    return validate_project_document(document)


def write_project(path: str | Path, document: Mapping[str, Any]) -> Path:
    """Validate and atomically replace a project file at a user-selected path."""
    target = normalise_project_path(path)
    validate_project_document(dict(document))
    if not target.parent.exists() or not target.parent.is_dir():
        raise ProjectError("项目保存位置不存在。")
    temporary_name: str | None = None
    try:
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{target.stem}-", suffix=".tmp", dir=target.parent,
        )
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(document, handle, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, target)
        temporary_name = None
    except (OSError, TypeError, ValueError) as exc:
        raise ProjectError("无法安全写入项目文件；原有项目未被替换。") from exc
    finally:
        if temporary_name:
            try:
                Path(temporary_name).unlink(missing_ok=True)
            except OSError:
                pass
    return target


def read_project(path: str | Path) -> dict[str, Any]:
    """Read a UTF-8 JSON project without executing any embedded content."""
    source = Path(path)
    try:
        with source.open("r", encoding="utf-8") as handle:
            document = json.load(
                handle,
                parse_constant=lambda value: (_ for _ in ()).throw(
                    ProjectError(f"项目文件包含不允许的 JSON 常量：{value}")
                ),
            )
    except ProjectError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError, RecursionError) as exc:
        raise ProjectError("无法读取项目文件，文件可能已损坏或不是 UTF-8 JSON。") from exc
    return validate_project_document(document)
