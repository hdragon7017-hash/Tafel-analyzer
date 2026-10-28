"""Durable, user-owned parameter templates kept outside analysis projects."""
from __future__ import annotations

import json
import math
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path


PRESET_FORMAT = "tafel-analyzer-parameter-templates"
PRESET_VERSION = 1
_MAX_TEMPLATES = 200
_MAX_NAME_LENGTH = 80
_UNSAFE_NAME_CHARS = set('<>:"/\\|?*')


class PresetError(ValueError):
    """A recoverable local-template storage or validation problem."""


@dataclass(frozen=True)
class ParameterTemplate:
    name: str
    area_cm2: float
    cathodic_range_v: tuple[float, float]
    anodic_range_v: tuple[float, float]
    export_range_v: tuple[float, float]


def default_preset_path() -> Path:
    """Return one fixed per-user LocalAppData location, never a project path."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    base = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
    return base / "TafelAnalyzer" / "parameter_templates.json"


def _number(value: object, field: str) -> float:
    if isinstance(value, bool):
        raise PresetError(f"模板字段“{field}”必须是有限数值。")
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise PresetError(f"模板字段“{field}”必须是有限数值。") from exc
    if not math.isfinite(number):
        raise PresetError(f"模板字段“{field}”必须是有限数值。")
    return number


def _range(value: object, field: str) -> tuple[float, float]:
    if not isinstance(value, (list, tuple)) or len(value) != 2:
        raise PresetError(f"模板字段“{field}”必须包含两个数值。")
    low, high = _number(value[0], field), _number(value[1], field)
    if low >= high:
        raise PresetError(f"模板字段“{field}”下限必须小于上限。")
    return low, high


def _name(value: object) -> str:
    if not isinstance(value, str):
        raise PresetError("模板名称必须是文本。")
    name = value.strip()
    if not name or len(name) > _MAX_NAME_LENGTH:
        raise PresetError(f"模板名称不能为空且不能超过 {_MAX_NAME_LENGTH} 个字符。")
    if any(ord(character) < 32 for character in name) or any(character in _UNSAFE_NAME_CHARS for character in name):
        raise PresetError("模板名称不能包含控制字符或文件路径特殊字符。")
    return name


def _template_from_object(value: object) -> ParameterTemplate:
    if not isinstance(value, dict):
        raise PresetError("模板记录格式无效。")
    area = _number(value.get("area_cm2"), "area_cm2")
    if area <= 0:
        raise PresetError("模板面积必须大于零。")
    return ParameterTemplate(
        _name(value.get("name")), area,
        _range(value.get("cathodic_range_v"), "cathodic_range_v"),
        _range(value.get("anodic_range_v"), "anodic_range_v"),
        _range(value.get("export_range_v"), "export_range_v"),
    )


class ParameterPresetStore:
    """Read/write one small atomic JSON document without touching .tafelproj."""

    def __init__(self, path: str | Path | None = None) -> None:
        self.path = Path(path) if path is not None else default_preset_path()

    def _load_document(self) -> tuple[list[ParameterTemplate], str | None]:
        if not self.path.exists():
            return [], None
        try:
            with self.path.open("r", encoding="utf-8") as handle:
                document = json.load(handle, parse_constant=lambda value: (_ for _ in ()).throw(PresetError(f"模板文件包含不允许的 JSON 常量：{value}")))
        except PresetError:
            raise
        except (OSError, UnicodeError, json.JSONDecodeError, RecursionError) as exc:
            raise PresetError("参数模板文件无法读取或已损坏；原文件已保留，未覆盖任何数据。") from exc
        if not isinstance(document, dict) or document.get("format") != PRESET_FORMAT or document.get("version") != PRESET_VERSION:
            raise PresetError("参数模板文件格式或版本无效；原文件已保留，未覆盖任何数据。")
        raw_templates = document.get("templates")
        if not isinstance(raw_templates, list) or len(raw_templates) > _MAX_TEMPLATES:
            raise PresetError("参数模板文件中的模板列表无效；原文件已保留，未覆盖任何数据。")
        templates = [_template_from_object(item) for item in raw_templates]
        names = [template.name.casefold() for template in templates]
        if len(set(names)) != len(names):
            raise PresetError("参数模板文件包含重复名称；原文件已保留，未覆盖任何数据。")
        startup = document.get("startup_default")
        if startup is not None:
            startup = _name(startup)
            if startup.casefold() not in names:
                raise PresetError("参数模板文件的启动默认模板不存在；原文件已保留，未覆盖任何数据。")
        return templates, startup

    @staticmethod
    def _document(templates: list[ParameterTemplate], startup: str | None) -> dict[str, object]:
        return {
            "format": PRESET_FORMAT,
            "version": PRESET_VERSION,
            "startup_default": startup,
            "templates": [
                {
                    "name": template.name,
                    "area_cm2": template.area_cm2,
                    "cathodic_range_v": list(template.cathodic_range_v),
                    "anodic_range_v": list(template.anodic_range_v),
                    "export_range_v": list(template.export_range_v),
                }
                for template in templates
            ],
        }

    def _write_document(self, templates: list[ParameterTemplate], startup: str | None) -> None:
        if len(templates) > _MAX_TEMPLATES:
            raise PresetError(f"最多保存 {_MAX_TEMPLATES} 个参数模板。")
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise PresetError("无法创建参数模板保存位置。") from exc
        temporary_name: str | None = None
        try:
            descriptor, temporary_name = tempfile.mkstemp(prefix=".parameter_templates-", suffix=".tmp", dir=self.path.parent)
            with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
                json.dump(self._document(templates, startup), handle, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary_name, self.path)
            temporary_name = None
        except (OSError, TypeError, ValueError) as exc:
            raise PresetError("无法安全写入参数模板；原有模板未被替换。") from exc
        finally:
            if temporary_name is not None:
                try:
                    Path(temporary_name).unlink(missing_ok=True)
                except OSError:
                    pass

    def list_templates(self) -> list[ParameterTemplate]:
        return self._load_document()[0]

    def startup_template(self) -> ParameterTemplate | None:
        templates, startup = self._load_document()
        if startup is None:
            return None
        return next(template for template in templates if template.name.casefold() == startup.casefold())

    def save(self, template: ParameterTemplate, *, overwrite: bool = False) -> ParameterTemplate:
        validated = _template_from_object({
            "name": template.name,
            "area_cm2": template.area_cm2,
            "cathodic_range_v": template.cathodic_range_v,
            "anodic_range_v": template.anodic_range_v,
            "export_range_v": template.export_range_v,
        })
        templates, startup = self._load_document()  # Never overwrite a corrupt document.
        index = next((index for index, item in enumerate(templates) if item.name.casefold() == validated.name.casefold()), None)
        if index is not None and not overwrite:
            raise PresetError(f"参数模板“{validated.name}”已存在。")
        if index is None:
            templates.append(validated)
        else:
            templates[index] = validated
            if startup is not None and startup.casefold() == validated.name.casefold():
                startup = validated.name
        self._write_document(templates, startup)
        return validated

    def delete(self, name: str) -> None:
        requested = _name(name)
        templates, startup = self._load_document()
        retained = [item for item in templates if item.name.casefold() != requested.casefold()]
        if len(retained) == len(templates):
            raise PresetError(f"参数模板“{requested}”不存在。")
        if startup is not None and startup.casefold() == requested.casefold():
            startup = None
        self._write_document(retained, startup)

    def set_startup_default(self, name: str | None) -> None:
        templates, _ = self._load_document()
        if name is None:
            self._write_document(templates, None)
            return
        requested = _name(name)
        selected = next((item for item in templates if item.name.casefold() == requested.casefold()), None)
        if selected is None:
            raise PresetError(f"参数模板“{requested}”不存在。")
        self._write_document(templates, selected.name)
