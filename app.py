#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import binascii
import errno
import hashlib
import json
import mimetypes
import os
import platform
import re
import shutil
import socket
import struct
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import datetime
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
GENERATED_DIR = ROOT / "generated"
KEYNOTE_EXPORT_SCRIPT = ROOT / "scripts" / "export_keynote.applescript"
POWERPOINT_EXPORT_SCRIPT = ROOT / "scripts" / "export_powerpoint.ps1"
WINDOWS_IMAGE_SCRIPT = ROOT / "scripts" / "convert_image_windows.ps1"
METADATA_FILENAME = "metadata.json"

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P_NS, "a": A_NS, "r": R_NS}
GENERATE_LOCK = threading.Lock()
FIXED_FOOTER_TEXT = "星芽铁军 战无不胜 攻无不克"

FONT_ALIASES_BY_SYSTEM = {
    "Darwin": {
        "微软雅黑": "PingFang SC",
        "黑体": "STHeiti",
        "宋体": "Songti SC",
        "楷体": "Kaiti SC",
        "苹方": "PingFang SC",
        "冬青黑体": "Hiragino Sans",
    },
    "Windows": {
        "微软雅黑": "Microsoft YaHei",
        "黑体": "SimHei",
        "宋体": "SimSun",
        "楷体": "KaiTi",
        "苹方": "Microsoft YaHei",
        "冬青黑体": "Microsoft YaHei",
    },
}


@dataclass(frozen=True)
class Field:
    key: str
    label: str
    shape_id: str
    default: str
    kind: str = "text"
    run_index: int | None = None
    hint: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "label": self.label,
            "default": self.default,
            "kind": self.kind,
            "hint": self.hint,
        }


BUILTIN_FIELDS: dict[str, list[dict[str, Any]]] = {
    "Q1战报模版-竖版.pptx": [
        {"key": "product", "label": "产品", "shape_id": "2", "run_index": 0},
        {"key": "project", "label": "项目", "shape_id": "2", "run_index": 2},
        {"key": "headline", "label": "竖排主标题", "shape_id": "7"},
        {"key": "name", "label": "姓名 / 战报名称", "shape_id": "8"},
        {"key": "company", "label": "公司名称", "shape_id": "9"},
        {"key": "team", "label": "团队", "shape_id": "27", "run_index": 1},
        {"key": "manager", "label": "主管", "shape_id": "28", "run_index": 1},
        {"key": "amount", "label": "新签收款", "shape_id": "18", "kind": "number"},
        {
            "key": "experience",
            "label": "经验分享",
            "shape_id": "23",
            "kind": "textarea",
        },
    ],
    "Q1战报模版-横版.pptx": [
        {"key": "nickname", "label": "代用名", "shape_id": "28"},
        {"key": "region", "label": "区域", "shape_id": "29"},
        {"key": "team", "label": "团队", "shape_id": "13", "run_index": 1},
        {"key": "manager", "label": "主管", "shape_id": "13", "run_index": 4},
        {"key": "amount", "label": "新签收款", "shape_id": "18", "kind": "number"},
        {
            "key": "experience",
            "label": "经验分享",
            "shape_id": "23",
            "kind": "textarea",
        },
    ],
    "4星保效方案.pptx": [
        {"key": "headline", "label": "竖排主标题", "shape_id": "10"},
        {"key": "name", "label": "姓名 / 战报名称", "shape_id": "11"},
        {"key": "company", "label": "公司名称", "shape_id": "13"},
        {"key": "team", "label": "团队", "shape_id": "27", "run_index": 1},
        {"key": "manager", "label": "主管", "shape_id": "28", "run_index": 1},
        {"key": "amount", "label": "新签收款", "shape_id": "18", "kind": "number"},
        {
            "key": "experience",
            "label": "经验分享",
            "shape_id": "23",
            "kind": "textarea",
        },
    ],
    "5星保效方案.pptx": [
        {"key": "headline", "label": "竖排主标题", "shape_id": "30"},
        {"key": "name", "label": "姓名 / 战报名称", "shape_id": "29"},
        {"key": "company", "label": "公司名称", "shape_id": "8"},
        {"key": "team", "label": "团队", "shape_id": "27", "run_index": 1},
        {"key": "manager", "label": "主管", "shape_id": "28", "run_index": 1},
        {"key": "amount", "label": "新签收款", "shape_id": "18", "kind": "number"},
        {
            "key": "experience",
            "label": "经验分享",
            "shape_id": "23",
            "kind": "textarea",
        },
    ],
    "A200方案.pptx": [
        {"key": "left_title", "label": "左侧竖排文字", "shape_id": "8"},
        {"key": "right_title", "label": "右侧竖排文字", "shape_id": "16"},
        {"key": "company", "label": "公司名称", "shape_id": "24"},
        {"key": "team", "label": "团队", "shape_id": "27", "run_index": 1},
        {"key": "manager", "label": "主管", "shape_id": "28", "run_index": 1},
        {"key": "amount", "label": "新签收款", "shape_id": "18", "kind": "number"},
        {
            "key": "experience",
            "label": "经验分享",
            "shape_id": "23",
            "kind": "textarea",
        },
    ],
    "Q1战报模版-竖版accio-okki.pptx": [
        {"key": "name", "label": "姓名 / 战报名称", "shape_id": "41"},
        {"key": "company", "label": "公司名称", "shape_id": "42"},
        {"key": "team", "label": "SAAS / 团队", "shape_id": "46"},
        {"key": "manager", "label": "KP / 主管", "shape_id": "62"},
        {"key": "product", "label": "签约产品", "shape_id": "63"},
        {"key": "amount", "label": "新签收款", "shape_id": "65", "kind": "number"},
        {"key": "headline", "label": "顶部标语", "shape_id": "37"},
        {"key": "footer", "label": "底部口号", "shape_id": "35", "fixed": True},
    ],
    "Q1战报模版-竖版accio.pptx": [
        {"key": "product", "label": "签约产品 / 版本", "shape_id": "10"},
        {"key": "amount", "label": "新签收款", "shape_id": "18", "kind": "number"},
        {
            "key": "experience",
            "label": "经验分享",
            "shape_id": "23",
            "kind": "textarea",
        },
        {"key": "name", "label": "姓名 / 战报名称", "shape_id": "29"},
        {"key": "company", "label": "公司名称", "shape_id": "3"},
        {"key": "team", "label": "团队", "shape_id": "4"},
        {"key": "manager", "label": "主管", "shape_id": "7"},
        {"key": "footer", "label": "底部口号", "shape_id": "8", "fixed": True},
    ],
    "Q1战报模版-竖版okki.pptx": [
        {"key": "name", "label": "姓名 / 战报名称", "shape_id": "18"},
        {"key": "company", "label": "公司名称", "shape_id": "29"},
        {"key": "team", "label": "SAAS / 团队", "shape_id": "10"},
        {"key": "kp", "label": "KP", "shape_id": "20"},
        {"key": "manager", "label": "主管", "shape_id": "14"},
        {"key": "product", "label": "签约产品", "shape_id": "2"},
        {"key": "amount", "label": "新签收款", "shape_id": "31", "kind": "number"},
        {"key": "headline", "label": "顶部标语", "shape_id": "36"},
        {"key": "footer", "label": "底部口号", "shape_id": "35", "fixed": True},
    ],
    "Q1战报模版-竖版okki5年.pptx": [
        {"key": "name", "label": "姓名 / 战报名称", "shape_id": "41"},
        {"key": "company", "label": "公司名称", "shape_id": "42"},
        {"key": "team", "label": "SAAS / 团队", "shape_id": "46"},
        {"key": "manager", "label": "KP / 主管", "shape_id": "62"},
        {"key": "product", "label": "签约产品", "shape_id": "63"},
        {"key": "amount", "label": "新签收款", "shape_id": "65", "kind": "number"},
        {"key": "headline", "label": "顶部标语", "shape_id": "37"},
        {"key": "footer", "label": "底部口号", "shape_id": "35", "fixed": True},
    ],
}

PERSON_IMAGE_BY_TEMPLATE = {
    "Q1战报模版-竖版.pptx": "ppt/media/image2.png",
    "Q1战报模版-横版.pptx": "ppt/media/image4.png",
    "Q1战报模版-竖版accio-okki.pptx": "ppt/media/image2.png",
    "Q1战报模版-竖版accio.pptx": "ppt/media/image2.png",
    "Q1战报模版-竖版okki.pptx": "ppt/media/image6.png",
    "Q1战报模版-竖版okki5年.pptx": "ppt/media/image8.png",
    "4星保效方案.pptx": "ppt/media/image2.png",
    "5星保效方案.pptx": "ppt/media/image2.png",
    "A200方案.pptx": "ppt/media/image2.png",
}

HORIZONTAL_GALLERY_SHAPES = ("15", "16", "17", "19", "22")
A200_SECOND_PERSON_PICTURE_ID = "10"


def _slot_aspect(template_path: Path | None, template_name: str, slot_key: str) -> dict[str, Any]:
    if template_path is None:
        return {"aspectRatio": 1, "targetWidth": 1000, "targetHeight": 1000}
    try:
        if slot_key.startswith("gallery"):
            index = int(slot_key.replace("gallery", ""))
            shape_id = HORIZONTAL_GALLERY_SHAPES[index - 1]
            width, height = _upload_pixel_size(_shape_box(template_path, shape_id))
        else:
            media_path = PERSON_IMAGE_BY_TEMPLATE.get(template_name)
            if not media_path:
                return {"aspectRatio": 1, "targetWidth": 1000, "targetHeight": 1000}
            width, height = _person_image_size(template_path, media_path)
        return {
            "aspectRatio": width / height if height else 1,
            "targetWidth": width,
            "targetHeight": height,
        }
    except Exception:
        return {"aspectRatio": 1, "targetWidth": 1000, "targetHeight": 1000}


def _upload_slot(
    key: str, label: str, hint: str, template_name: str, template_path: Path | None
) -> dict[str, Any]:
    return {
        "key": key,
        "label": label,
        "hint": hint,
        **_slot_aspect(template_path, template_name, key),
    }


def upload_slots_for_template(
    template_name: str, template_path: Path | None = None
) -> list[dict[str, Any]]:
    if template_name == "A200方案.pptx":
        return [
            _upload_slot("person", "左侧人物照片", "上传左侧人物照片", template_name, template_path),
            _upload_slot("person2", "右侧人物照片", "上传右侧人物照片", template_name, template_path),
        ]
    slots = [_upload_slot("person", "人物照片", "上传人物照片", template_name, template_path)]
    if template_name == "Q1战报模版-横版.pptx":
        slots.extend(
            _upload_slot(
                f"gallery{index}",
                f"战报图片 {index}",
                f"上传白框图片 {index}",
                template_name,
                template_path,
            )
            for index in range(1, 6)
        )
    return slots

STATIC_TEXTS = {
    "新签收款",
    "经验分享",
    "经验",
    "分享",
    "战绩",
    "团队：",
    "主管：",
    "主管",
    "：",
    "+",
}


def _natural_slide_key(name: str) -> tuple[int, str]:
    match = re.search(r"slide(\d+)\.xml$", name)
    return (int(match.group(1)) if match else 0, name)


def _register_namespaces(xml_bytes: bytes) -> None:
    for _, value in ET.iterparse(
        _BytesReader(xml_bytes), events=("start-ns",)
    ):
        prefix, uri = value
        ET.register_namespace(prefix or "", uri)


class _BytesReader:
    def __init__(self, value: bytes) -> None:
        self.value = value
        self.offset = 0

    def read(self, size: int = -1) -> bytes:
        if size < 0:
            size = len(self.value) - self.offset
        result = self.value[self.offset : self.offset + size]
        self.offset += size
        return result


def _slide_size(archive: zipfile.ZipFile) -> tuple[int, int]:
    root = ET.fromstring(archive.read("ppt/presentation.xml"))
    node = root.find("p:sldSz", NS)
    if node is None:
        return 0, 0
    return int(node.get("cx", "0")), int(node.get("cy", "0"))


def _slide_names(archive: zipfile.ZipFile) -> list[str]:
    return sorted(
        [
            name
            for name in archive.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
        ],
        key=_natural_slide_key,
    )


def _shape_text_nodes(shape: ET.Element) -> list[ET.Element]:
    return shape.findall(".//a:t", NS)


def _shape_id(shape: ET.Element) -> str:
    node = shape.find("./p:nvSpPr/p:cNvPr", NS)
    return node.get("id", "") if node is not None else ""


def _is_visible(shape: ET.Element, slide_width: int, slide_height: int) -> bool:
    transform = shape.find("./p:spPr/a:xfrm", NS)
    if transform is None:
        return True
    offset = transform.find("a:off", NS)
    extent = transform.find("a:ext", NS)
    if offset is None or extent is None:
        return True
    x, y = int(offset.get("x", "0")), int(offset.get("y", "0"))
    width, height = int(extent.get("cx", "0")), int(extent.get("cy", "0"))
    return x < slide_width and y < slide_height and x + width > 0 and y + height > 0


def _shape_map(template_path: Path) -> tuple[dict[str, list[str]], tuple[int, int]]:
    with zipfile.ZipFile(template_path) as archive:
        width, height = _slide_size(archive)
        slides = _slide_names(archive)
        if not slides:
            return {}, (width, height)
        root = ET.fromstring(archive.read(slides[0]))
        shapes: dict[str, list[str]] = {}
        for shape in root.findall(".//p:sp", NS):
            if not _is_visible(shape, width, height):
                continue
            texts = [(node.text or "") for node in _shape_text_nodes(shape)]
            if texts:
                shapes[_shape_id(shape)] = texts
        return shapes, (width, height)


def _field_default(shape_texts: list[str], run_index: int | None) -> str:
    if run_index is None:
        return "".join(shape_texts)
    return shape_texts[run_index] if 0 <= run_index < len(shape_texts) else ""


def _generic_fields(
    shapes: dict[str, list[str]], excluded_shape_ids: set[str] | None = None
) -> list[Field]:
    excluded_shape_ids = excluded_shape_ids or set()
    fields: list[Field] = []
    for shape_id, texts in shapes.items():
        if shape_id in excluded_shape_ids:
            continue
        value = "".join(texts).strip()
        if not value or value in STATIC_TEXTS:
            continue
        kind = "textarea" if len(value) > 30 else "text"
        fields.append(
            Field(
                key=f"shape_{shape_id}",
                label=f"文本框 {shape_id}",
                shape_id=shape_id,
                default=value,
                kind=kind,
                hint="新版模板中自动识别的文本框",
            )
        )
    return fields


def fields_for_template(template_path: Path) -> list[Field]:
    shapes, _ = _shape_map(template_path)
    configured = BUILTIN_FIELDS.get(template_path.name)
    if configured:
        result: list[Field] = []
        configured_shape_ids: set[str] = set()
        for item in configured:
            shape_id = str(item["shape_id"])
            texts = shapes.get(shape_id, [])
            run_index = item.get("run_index")
            if not texts or (run_index is not None and run_index >= len(texts)):
                continue
            configured_shape_ids.add(shape_id)
            if item.get("fixed"):
                continue
            result.append(
                Field(
                    key=item["key"],
                    label=item["label"],
                    shape_id=shape_id,
                    default=_field_default(texts, run_index),
                    kind=item.get("kind", "text"),
                    run_index=run_index,
                    hint=item.get("hint", ""),
                )
            )
        return result + _generic_fields(shapes, configured_shape_ids)

    return _generic_fields(shapes)


def _template_error(path: Path, message: str) -> dict[str, str]:
    return {"id": path.name, "error": message}


def _is_template_candidate(path: Path) -> bool:
    name = path.name
    if path.suffix.lower() != ".pptx":
        return False
    return not (name.startswith(".") or name.startswith("~$") or name.startswith(".~"))


def _template_paths() -> list[Path]:
    return sorted(
        [path for path in ROOT.glob("*.pptx") if _is_template_candidate(path)],
        key=lambda item: item.name,
    )


def _template_item(path: Path) -> dict[str, Any]:
    _, (width, height) = _shape_map(path)
    orientation = "横版" if width > height else "竖版"
    fields = fields_for_template(path)
    return {
        "id": path.name,
        "name": path.stem,
        "orientation": orientation,
        "fieldCount": len(fields),
        "fields": [field.as_dict() for field in fields],
        "uploadSlots": upload_slots_for_template(path.name, path),
        "builtin": path.name in BUILTIN_FIELDS,
    }


def templates_with_errors() -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    items: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []
    for path in _template_paths():
        try:
            items.append(_template_item(path))
        except zipfile.BadZipFile:
            errors.append(_template_error(path, "不是有效的 PPTX 文件，已跳过"))
        except (KeyError, ET.ParseError, OSError, ValueError) as exc:
            errors.append(_template_error(path, f"读取失败：{exc}"))
    return items, errors


def templates() -> list[dict[str, Any]]:
    return templates_with_errors()[0]


def templates_payload() -> dict[str, Any]:
    paths = _template_paths()
    digest = hashlib.sha256()
    latest_modified = 0.0
    for path in paths:
        stat = path.stat()
        latest_modified = max(latest_modified, stat.st_mtime)
        digest.update(path.name.encode("utf-8"))
        digest.update(str(stat.st_size).encode("ascii"))
        digest.update(str(stat.st_mtime_ns).encode("ascii"))
    template_items, template_errors = templates_with_errors()
    return {
        "templates": template_items,
        "templateErrors": template_errors,
        "revision": digest.hexdigest()[:12],
        "refreshedAt": datetime.now().isoformat(timespec="seconds"),
        "latestModifiedAt": (
            datetime.fromtimestamp(latest_modified).isoformat(timespec="seconds")
            if latest_modified
            else ""
        ),
    }


def _template_path(template_name: str) -> Path:
    name = Path(template_name).name
    path = ROOT / name
    if name != template_name or path.suffix.lower() != ".pptx" or not path.exists():
        raise ValueError("找不到选择的 PPTX 模板")
    return path


def _replace_shape_text(
    root: ET.Element,
    shape_id: str,
    run_index: int | None,
    value: str,
    fit_amount_from: str = "",
) -> None:
    for shape in root.findall(".//p:sp", NS):
        if _shape_id(shape) != shape_id:
            continue
        nodes = _shape_text_nodes(shape)
        if not nodes:
            return
        if run_index is None:
            nodes[0].text = value
            for node in nodes[1:]:
                node.text = ""
        elif run_index < len(nodes):
            nodes[run_index].text = value
        if fit_amount_from:
            _center_shape_text(shape)
            _fit_shape_font(
                shape,
                fit_amount_from,
                value,
                min_size=2200,
                padding=0.74,
                always_fit=True,
            )
        return


def _center_shape_text(shape: ET.Element) -> None:
    for paragraph in shape.findall(".//a:p", NS):
        properties = paragraph.find("a:pPr", NS)
        if properties is None:
            properties = ET.Element(f"{{{A_NS}}}pPr")
            paragraph.insert(0, properties)
        properties.set("algn", "ctr")


def _text_width(value: str) -> float:
    digit_widths = {
        "0": 0.82,
        "1": 0.48,
        "2": 0.86,
        "3": 0.86,
        "4": 0.92,
        "5": 0.86,
        "6": 0.92,
        "7": 0.82,
        "8": 1.0,
        "9": 0.92,
        ",": 0.28,
        ".": 0.28,
        " ": 0.3,
    }
    return sum(digit_widths.get(character, 1.0) for character in value)


def _fit_shape_font(
    shape: ET.Element,
    placeholder: str,
    value: str,
    min_size: int = 12000,
    padding: float = 0.96,
    always_fit: bool = False,
) -> None:
    placeholder_width = _text_width(placeholder)
    value_width = _text_width(value)
    if not placeholder_width or (value_width <= placeholder_width and not always_fit):
        return
    scale = min(1.0, placeholder_width / value_width * padding)
    for properties in shape.findall(".//a:rPr", NS) + shape.findall(".//a:endParaRPr", NS):
        size = properties.get("sz")
        if size:
            original_size = int(size)
            fitted_size = max(min_size, round(original_size * scale))
            properties.set("sz", str(min(original_size, fitted_size)))


def _fit_shape_from_current_text(root: ET.Element, shape_id: str, placeholder: str) -> None:
    for shape in root.findall(".//p:sp", NS):
        if _shape_id(shape) != shape_id:
            continue
        value = "".join(node.text or "" for node in _shape_text_nodes(shape))
        _fit_shape_font(shape, placeholder, value)
        return


def _a_tag(name: str) -> str:
    return f"{{{A_NS}}}{name}"


def _insert_after_text_fill(properties: ET.Element, child: ET.Element) -> None:
    fill_tags = {
        _a_tag("noFill"),
        _a_tag("solidFill"),
        _a_tag("gradFill"),
        _a_tag("blipFill"),
        _a_tag("pattFill"),
        _a_tag("grpFill"),
    }
    children = list(properties)
    index = 1 if children and children[0].tag == _a_tag("ln") else 0
    while index < len(children) and children[index].tag in fill_tags:
        index += 1
    properties.insert(index, child)


def _set_text_outline(
    properties: ET.Element, color: str, width: int, alpha: int = 100000
) -> None:
    for child in list(properties):
        if child.tag == _a_tag("ln"):
            properties.remove(child)
    line = ET.Element(_a_tag("ln"), {"w": str(width), "cap": "rnd"})
    fill = ET.SubElement(line, _a_tag("solidFill"))
    rgb = ET.SubElement(fill, _a_tag("srgbClr"), {"val": color})
    if alpha < 100000:
        ET.SubElement(rgb, _a_tag("alpha"), {"val": str(alpha)})
    properties.insert(0, line)


def _set_text_shadow(
    properties: ET.Element,
    color: str,
    alpha: int,
    blur: int,
    distance: int,
    direction: int = 2700000,
) -> None:
    for child in list(properties):
        if child.tag == _a_tag("effectLst"):
            properties.remove(child)
    effects = ET.Element(_a_tag("effectLst"))
    shadow = ET.SubElement(
        effects,
        _a_tag("outerShdw"),
        {
            "blurRad": str(blur),
            "dist": str(distance),
            "dir": str(direction),
            "algn": "ctr",
            "rotWithShape": "0",
        },
    )
    rgb = ET.SubElement(shadow, _a_tag("srgbClr"), {"val": color})
    ET.SubElement(rgb, _a_tag("alpha"), {"val": str(alpha)})
    _insert_after_text_fill(properties, effects)


def _set_text_solid_fill(properties: ET.Element, color: str) -> None:
    for child in list(properties):
        if child.tag in {_a_tag("solidFill"), _a_tag("gradFill")}:
            properties.remove(child)
    fill = ET.Element(_a_tag("solidFill"))
    ET.SubElement(fill, _a_tag("srgbClr"), {"val": color})
    index = 1 if list(properties) and list(properties)[0].tag == _a_tag("ln") else 0
    properties.insert(index, fill)


def _clean_text_color(value: Any) -> str:
    color = str(value or "").strip().lstrip("#").upper()
    return color if re.fullmatch(r"[0-9A-F]{6}", color) else ""


def _clean_font_family(value: Any) -> str:
    font = re.sub(r"[\x00-\x1f<>]", "", str(value or "").strip())
    if not font:
        return ""
    aliases = FONT_ALIASES_BY_SYSTEM.get(platform.system(), {})
    return aliases.get(font, font)[:64]


def _clean_font_size(value: Any) -> int | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    try:
        points = float(raw)
    except ValueError:
        return None
    points = min(300.0, max(6.0, points))
    return int(round(points * 100))


def _clean_bold(value: Any) -> str:
    raw = str(value).strip().lower()
    if raw in {"1", "true", "yes", "bold"}:
        return "1"
    if raw in {"0", "false", "no", "normal"}:
        return "0"
    return ""


def _set_typeface(properties: ET.Element, font: str) -> None:
    for tag in ("latin", "ea", "cs"):
        node = properties.find(f"a:{tag}", NS)
        if node is None:
            node = ET.SubElement(properties, _a_tag(tag))
        node.set("typeface", font)


def _apply_custom_text_style(properties: ET.Element, style: Any) -> None:
    if not isinstance(style, dict):
        return
    color = _clean_text_color(style.get("color"))
    if color:
        _set_text_solid_fill(properties, color)
    font = _clean_font_family(style.get("font"))
    if font:
        _set_typeface(properties, font)
    size = _clean_font_size(style.get("size"))
    if size:
        properties.set("sz", str(size))
    bold = _clean_bold(style.get("bold"))
    if bold:
        properties.set("b", bold)


def _text_style_has_values(style: Any) -> bool:
    if not isinstance(style, dict):
        return False
    return any(str(style.get(key, "")).strip() for key in ("color", "font", "size", "bold"))


def _apply_custom_text_styles(
    root: ET.Element, fields: list[Field], text_styles: dict[str, Any]
) -> None:
    if not isinstance(text_styles, dict):
        return
    global_style = text_styles.get("__all")
    if _text_style_has_values(global_style):
        for properties in root.findall(".//a:rPr", NS) + root.findall(".//a:endParaRPr", NS):
            _apply_custom_text_style(properties, global_style)
    for field in fields:
        style = text_styles.get(field.key)
        if not _text_style_has_values(style):
            continue
        shape = _shape_by_id(root, field.shape_id)
        if shape is None:
            continue
        for properties in _shape_properties(shape, field.run_index):
            _apply_custom_text_style(properties, style)


def _apply_text_style(properties: ET.Element, style: str) -> None:
    if style == "poster_title":
        properties.set("b", "1")
        properties.set("spc", "20")
        _set_text_outline(properties, "7E3522", 5200, 76000)
        _set_text_shadow(properties, "4A160E", 34000, 15000, 7500)
    elif style == "amount":
        properties.set("b", "1")
        _set_text_outline(properties, "18225E", 5200, 62000)
        _set_text_shadow(properties, "0B123E", 30000, 12000, 5200)


def _shape_by_id(root: ET.Element, shape_id: str) -> ET.Element | None:
    for shape in root.findall(".//p:sp", NS):
        if _shape_id(shape) == shape_id:
            return shape
    return None


def _shape_properties(
    shape: ET.Element, run_index: int | None = None
) -> list[ET.Element]:
    if run_index is None:
        return shape.findall(".//a:rPr", NS) + shape.findall(".//a:endParaRPr", NS)
    runs = shape.findall(".//a:r", NS)
    if not 0 <= run_index < len(runs):
        return []
    properties = runs[run_index].find("a:rPr", NS)
    return [properties] if properties is not None else []


def _fit_field_font(
    shape: ET.Element,
    run_index: int | None,
    value: str,
    capacity: int,
    min_size: int = 2800,
) -> None:
    value_width = _text_width(value)
    if value_width <= capacity:
        return
    scale = capacity / value_width * 0.96
    for properties in _shape_properties(shape, run_index):
        size = properties.get("sz")
        if size:
            original_size = int(size)
            readable_floor = min(min_size, max(700, round(original_size * 0.58)))
            fitted_size = max(readable_floor, round(original_size * scale))
            properties.set("sz", str(min(original_size, fitted_size)))


def _style_dynamic_fields(root: ET.Element, fields: list[Field]) -> None:
    style_by_key = {
        "amount": "amount",
    }
    capacity_by_key = {
        "headline": 7,
        "name": 7,
        "left_title": 7,
        "right_title": 7,
        "nickname": 8,
        "region": 8,
        "company": 13,
        "team": 7,
        "manager": 7,
        "experience": 52,
    }
    styled_shapes: set[tuple[str, str]] = set()
    for field in fields:
        shape = _shape_by_id(root, field.shape_id)
        if shape is None:
            continue
        style = style_by_key.get(field.key)
        if style:
            properties = _shape_properties(shape, field.run_index)
            style_key = (field.shape_id, style)
            if style_key not in styled_shapes:
                for item in properties:
                    _apply_text_style(item, style)
                styled_shapes.add(style_key)
        value = _field_default([node.text or "" for node in _shape_text_nodes(shape)], field.run_index)
        capacity = capacity_by_key.get(field.key)
        if capacity:
            _fit_field_font(shape, field.run_index, value, capacity)


def _p_tag(name: str) -> str:
    return f"{{{P_NS}}}{name}"


def _r_tag(name: str) -> str:
    return f"{{{RELATIONSHIPS_NS}}}{name}"


def _next_shape_id(root: ET.Element) -> str:
    shape_ids = []
    for node in root.findall(".//p:cNvPr", NS):
        value = node.get("id", "")
        if value.isdigit():
            shape_ids.append(int(value))
    return str(max(shape_ids, default=0) + 1)


def _next_relationship_id(root: ET.Element) -> str:
    relationship_ids = []
    for node in root.findall(_r_tag("Relationship")):
        value = node.get("Id", "")
        match = re.fullmatch(r"rId(\d+)", value)
        if match:
            relationship_ids.append(int(match.group(1)))
    return f"rId{max(relationship_ids, default=0) + 1}"


def _add_image_relationship(relationships: ET.Element, media_path: str) -> str:
    relationship_id = _next_relationship_id(relationships)
    ET.SubElement(
        relationships,
        _r_tag("Relationship"),
        {
            "Id": relationship_id,
            "Type": f"{R_NS}/image",
            "Target": f"../media/{Path(media_path).name}",
        },
    )
    return relationship_id


def _picture_transform(shape: ET.Element) -> tuple[int, int, int, int] | None:
    transform = shape.find("./p:spPr/a:xfrm", NS)
    if transform is None:
        return None
    offset = transform.find("a:off", NS)
    extent = transform.find("a:ext", NS)
    if offset is None or extent is None:
        return None
    return (
        int(offset.get("x", "0")),
        int(offset.get("y", "0")),
        int(extent.get("cx", "0")),
        int(extent.get("cy", "0")),
    )


def _group_transform(group: ET.Element) -> tuple[int, int, int, int, int, int, int, int] | None:
    transform = group.find("./p:grpSpPr/a:xfrm", NS)
    if transform is None:
        return None
    offset = transform.find("a:off", NS)
    extent = transform.find("a:ext", NS)
    child_offset = transform.find("a:chOff", NS)
    child_extent = transform.find("a:chExt", NS)
    if offset is None or extent is None or child_offset is None or child_extent is None:
        return None
    return (
        int(offset.get("x", "0")),
        int(offset.get("y", "0")),
        int(extent.get("cx", "0")),
        int(extent.get("cy", "0")),
        int(child_offset.get("x", "0")),
        int(child_offset.get("y", "0")),
        int(child_extent.get("cx", "0")),
        int(child_extent.get("cy", "0")),
    )


def _apply_group_transform(
    box: tuple[int, int, int, int],
    group: tuple[int, int, int, int, int, int, int, int],
) -> tuple[int, int, int, int]:
    x, y, width, height = box
    group_x, group_y, group_width, group_height, child_x, child_y, child_width, child_height = group
    scale_x = group_width / child_width if child_width else 1
    scale_y = group_height / child_height if child_height else 1
    return (
        round(group_x + (x - child_x) * scale_x),
        round(group_y + (y - child_y) * scale_y),
        round(width * scale_x),
        round(height * scale_y),
    )


def _shape_absolute_box_by_id(root: ET.Element, shape_id: str) -> tuple[int, int, int, int] | None:
    def walk(
        node: ET.Element,
        groups: list[tuple[int, int, int, int, int, int, int, int]],
    ) -> tuple[int, int, int, int] | None:
        for child in list(node):
            if child.tag == _p_tag("sp") and _shape_id(child) == shape_id:
                box = _picture_transform(child)
                if box is None:
                    return None
                for group in reversed(groups):
                    box = _apply_group_transform(box, group)
                return box
            if child.tag == _p_tag("grpSp"):
                group = _group_transform(child)
                result = walk(child, groups + ([group] if group else []))
                if result is not None:
                    return result
            elif child.tag != _p_tag("sp"):
                result = walk(child, groups)
                if result is not None:
                    return result
        return None

    return walk(root, [])


def _append_picture(
    root: ET.Element,
    relationships: ET.Element,
    media_path: str,
    box: tuple[int, int, int, int],
    name: str,
) -> None:
    shape_tree = root.find(".//p:spTree", NS)
    if shape_tree is None:
        return
    relationship_id = _add_image_relationship(relationships, media_path)
    x, y, width, height = box
    picture = ET.SubElement(shape_tree, _p_tag("pic"))
    non_visual = ET.SubElement(picture, _p_tag("nvPicPr"))
    ET.SubElement(
        non_visual,
        _p_tag("cNvPr"),
        {"id": _next_shape_id(root), "name": name},
    )
    locks = ET.SubElement(non_visual, _p_tag("cNvPicPr"))
    ET.SubElement(locks, _a_tag("picLocks"), {"noChangeAspect": "1"})
    ET.SubElement(non_visual, _p_tag("nvPr"))
    fill = ET.SubElement(picture, _p_tag("blipFill"))
    ET.SubElement(fill, _a_tag("blip"), {f"{{{R_NS}}}embed": relationship_id})
    stretch = ET.SubElement(fill, _a_tag("stretch"))
    ET.SubElement(stretch, _a_tag("fillRect"))
    properties = ET.SubElement(picture, _p_tag("spPr"))
    transform = ET.SubElement(properties, _a_tag("xfrm"))
    ET.SubElement(transform, _a_tag("off"), {"x": str(x), "y": str(y)})
    ET.SubElement(transform, _a_tag("ext"), {"cx": str(width), "cy": str(height)})
    geometry = ET.SubElement(properties, _a_tag("prstGeom"), {"prst": "rect"})
    ET.SubElement(geometry, _a_tag("avLst"))


def _relink_picture(
    root: ET.Element,
    relationships: ET.Element,
    picture_id: str,
    media_path: str,
) -> None:
    for picture in root.findall(".//p:pic", NS):
        properties = picture.find("./p:nvPicPr/p:cNvPr", NS)
        if properties is None or properties.get("id") != picture_id:
            continue
        blip = picture.find(".//a:blip", NS)
        if blip is not None:
            blip.set(f"{{{R_NS}}}embed", _add_image_relationship(relationships, media_path))
        return


def _append_gallery_images(
    root: ET.Element,
    relationships: ET.Element,
    gallery_media: list[tuple[str, str]],
) -> None:
    for shape_id, media_path in gallery_media:
        shape = _shape_by_id(root, shape_id)
        box = _picture_transform(shape) if shape is not None else None
        if box:
            _append_picture(root, relationships, media_path, box, f"上传战报图片 {shape_id}")


def _slide_contains_footer_slogan(root: ET.Element) -> bool:
    text = "".join(node.text or "" for node in root.findall(".//a:t", NS))
    return "星芽铁军" in text and ("战无不胜" in text or "攻无不克" in text)


def _fixed_footer_shape_id(template_name: str) -> str:
    for item in BUILTIN_FIELDS.get(template_name, []):
        if item.get("key") == "footer" and item.get("fixed"):
            return str(item["shape_id"])
    return ""


def _apply_fixed_footer_slogan(root: ET.Element, template_name: str) -> None:
    shape_id = _fixed_footer_shape_id(template_name)
    if shape_id:
        _replace_shape_text(root, shape_id, None, FIXED_FOOTER_TEXT)


def _append_footer_slogan(
    root: ET.Element, slide_width: int, slide_height: int
) -> None:
    shape_tree = root.find(".//p:spTree", NS)
    if shape_tree is None or not slide_width or not slide_height:
        return
    if slide_width > slide_height:
        x, y, width, height, font_size = 500000, 6080000, 11192000, 760000, 2400
    else:
        x, y, width, height, font_size = 650000, 35780000, 14474670, 1500000, 3900

    shape = ET.SubElement(shape_tree, _p_tag("sp"))
    non_visual = ET.SubElement(shape, _p_tag("nvSpPr"))
    ET.SubElement(
        non_visual,
        _p_tag("cNvPr"),
        {"id": _next_shape_id(root), "name": "固定底部口号"},
    )
    ET.SubElement(non_visual, _p_tag("cNvSpPr"), {"txBox": "1"})
    ET.SubElement(non_visual, _p_tag("nvPr"))

    shape_properties = ET.SubElement(shape, _p_tag("spPr"))
    transform = ET.SubElement(shape_properties, _a_tag("xfrm"))
    ET.SubElement(transform, _a_tag("off"), {"x": str(x), "y": str(y)})
    ET.SubElement(transform, _a_tag("ext"), {"cx": str(width), "cy": str(height)})
    geometry = ET.SubElement(shape_properties, _a_tag("prstGeom"), {"prst": "rect"})
    ET.SubElement(geometry, _a_tag("avLst"))
    ET.SubElement(shape_properties, _a_tag("noFill"))
    line = ET.SubElement(shape_properties, _a_tag("ln"))
    ET.SubElement(line, _a_tag("noFill"))

    text_body = ET.SubElement(shape, _p_tag("txBody"))
    ET.SubElement(
        text_body,
        _a_tag("bodyPr"),
        {"wrap": "none", "anchor": "ctr", "anchorCtr": "1"},
    )
    ET.SubElement(text_body, _a_tag("lstStyle"))
    paragraph = ET.SubElement(text_body, _a_tag("p"))
    ET.SubElement(paragraph, _a_tag("pPr"), {"algn": "ctr", "fontAlgn": "ctr"})
    run = ET.SubElement(paragraph, _a_tag("r"))
    properties = ET.SubElement(
        run,
        _a_tag("rPr"),
        {"lang": "zh-CN", "sz": str(font_size), "b": "1", "spc": "100"},
    )
    _set_text_solid_fill(properties, "FFD78A")
    _set_text_outline(properties, "7E3522", 3600, 62000)
    _set_text_shadow(properties, "4A160E", 26000, 10000, 5000)
    ET.SubElement(properties, _a_tag("latin"), {"typeface": "微软雅黑"})
    ET.SubElement(properties, _a_tag("ea"), {"typeface": "微软雅黑"})
    ET.SubElement(run, _a_tag("t")).text = FIXED_FOOTER_TEXT
    end_properties = ET.SubElement(
        paragraph,
        _a_tag("endParaRPr"),
        {"lang": "zh-CN", "sz": str(font_size), "b": "1"},
    )
    ET.SubElement(end_properties, _a_tag("latin"), {"typeface": "微软雅黑"})
    ET.SubElement(end_properties, _a_tag("ea"), {"typeface": "微软雅黑"})


def _style_product_project_title(root: ET.Element) -> None:
    shape = _shape_by_id(root, "2")
    if shape is None:
        return
    _fit_shape_font(shape, "产品+项目", "".join(node.text or "" for node in _shape_text_nodes(shape)))
    for properties in _shape_properties(shape):
        _apply_text_style(properties, "poster_title")
    runs = shape.findall(".//a:r", NS)
    if len(runs) >= 3:
        plus_properties = runs[1].find("a:rPr", NS)
        if plus_properties is not None:
            size = plus_properties.get("sz")
            if size:
                plus_properties.set("sz", str(max(10500, round(int(size) * 0.72))))
            _set_text_solid_fill(plus_properties, "FFD37A")


def render_pptx(
    template_path: Path,
    values: dict[str, str],
    output_path: Path,
    media_replacements: dict[str, bytes] | None = None,
    gallery_media: list[tuple[str, str]] | None = None,
    relink_images: list[tuple[str, str]] | None = None,
    text_styles: dict[str, Any] | None = None,
) -> None:
    fields = fields_for_template(template_path)
    media_replacements = media_replacements or {}
    gallery_media = gallery_media or []
    relink_images = relink_images or []
    text_styles = text_styles or {}
    with zipfile.ZipFile(template_path, "r") as source:
        slide_width, slide_height = _slide_size(source)
        slide_names = _slide_names(source)
        slide_name = slide_names[0] if slide_names else ""
        slide_rels_name = (
            f"ppt/slides/_rels/{Path(slide_name).name}.rels" if slide_name else ""
        )
        slide_xml = source.read(slide_name) if slide_name else b""
        slide_rels_xml = source.read(slide_rels_name) if slide_rels_name else b""
        if slide_xml:
            _register_namespaces(slide_xml)
            root = ET.fromstring(slide_xml)
            relationships = ET.fromstring(slide_rels_xml)
            for field in fields:
                _replace_shape_text(
                    root,
                    field.shape_id,
                    field.run_index,
                    str(values.get(field.key, field.default)),
                    field.default if field.key == "amount" else "",
                )
            if template_path.name == "Q1战报模版-竖版.pptx":
                _style_product_project_title(root)
            _style_dynamic_fields(root, fields)
            for picture_id, media_path in relink_images:
                _relink_picture(root, relationships, picture_id, media_path)
            _append_gallery_images(root, relationships, gallery_media)
            _apply_fixed_footer_slogan(root, template_path.name)
            if not _slide_contains_footer_slogan(root):
                _append_footer_slogan(root, slide_width, slide_height)
            _apply_custom_text_styles(root, fields, text_styles)
            slide_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            ET.register_namespace("", RELATIONSHIPS_NS)
            slide_rels_xml = ET.tostring(
                relationships, encoding="utf-8", xml_declaration=True
            )

        with zipfile.ZipFile(output_path, "w") as target:
            source_names = set(source.namelist())
            for item in source.infolist():
                if item.filename == slide_name:
                    data = slide_xml
                elif item.filename == slide_rels_name:
                    data = slide_rels_xml
                elif item.filename in media_replacements:
                    data = media_replacements[item.filename]
                else:
                    data = source.read(item.filename)
                target.writestr(item, data)
            for media_path, data in media_replacements.items():
                if media_path not in source_names:
                    target.writestr(media_path, data)


def _export_png_with_keynote(
    pptx_path: Path, render_dir: Path
) -> tuple[Path | None, str]:
    if not KEYNOTE_EXPORT_SCRIPT.exists():
        return None, "缺少 Keynote 导出脚本，已生成 PPTX。"
    if not Path("/Applications/Keynote.app").exists():
        return None, "本机未安装 Keynote，已生成 PPTX；请手动导出图片。"
    render_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        ["/usr/bin/open", "-gj", "-a", "Keynote"],
        check=False,
        capture_output=True,
        text=True,
        timeout=10,
    )
    try:
        completed = subprocess.run(
            [
                "/usr/bin/osascript",
                str(KEYNOTE_EXPORT_SCRIPT),
                str(pptx_path),
                str(render_dir),
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=75,
        )
    except subprocess.TimeoutExpired:
        return None, "Keynote 导出超时，PPTX 已保留。"
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()
        return None, f"Keynote 暂未导出图片：{detail or '未知错误'}"
    images = sorted(path for path in render_dir.rglob("*") if path.suffix.lower() == ".png")
    if not images:
        return None, "Keynote 已运行，但没有找到导出的 PNG；PPTX 已保留。"
    result = pptx_path.with_suffix(".png")
    shutil.copyfile(images[0], result)
    return result, ""


def _export_png_with_quicklook(
    pptx_path: Path, render_dir: Path
) -> tuple[Path | None, str]:
    quicklook = Path("/usr/bin/qlmanage")
    if not quicklook.exists():
        return None, "系统快速预览组件不可用。"
    quicklook_dir = render_dir / "quicklook"
    quicklook_dir.mkdir(parents=True, exist_ok=True)
    try:
        completed = subprocess.run(
            [
                str(quicklook),
                "-t",
                "-s",
                "2400",
                "-o",
                str(quicklook_dir),
                str(pptx_path),
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=45,
        )
    except subprocess.TimeoutExpired:
        return None, "系统快速预览导出超时。"
    images = sorted(path for path in quicklook_dir.glob("*.png") if path.is_file())
    if completed.returncode != 0 or not images:
        detail = (completed.stderr or completed.stdout).strip()
        return None, f"系统快速预览未能生成图片：{detail or '未知错误'}"
    result = pptx_path.with_suffix(".png")
    shutil.copyfile(images[0], result)
    return result, ""


def _powershell_executable() -> str | None:
    return shutil.which("powershell.exe") or shutil.which("powershell")


def _export_png_with_powerpoint(
    pptx_path: Path, render_dir: Path
) -> tuple[Path | None, str]:
    powershell = _powershell_executable()
    if not POWERPOINT_EXPORT_SCRIPT.exists():
        return None, "缺少 PowerPoint 导出脚本，已生成 PPTX。"
    if not powershell:
        return None, "未找到 Windows PowerShell，已生成 PPTX；请手动导出图片。"
    render_dir.mkdir(parents=True, exist_ok=True)
    try:
        completed = subprocess.run(
            [
                powershell,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(POWERPOINT_EXPORT_SCRIPT),
                "-SourcePath",
                str(pptx_path),
                "-TargetDir",
                str(render_dir),
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=150,
        )
    except subprocess.TimeoutExpired:
        return None, "PowerPoint 导出超时，PPTX 已保留。"
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()
        return None, f"PowerPoint 暂未导出图片：{detail or '请确认已安装桌面版 PowerPoint'}"
    images = sorted(path for path in render_dir.rglob("*") if path.suffix.lower() == ".png")
    if not images:
        return None, "PowerPoint 已运行，但没有找到导出的 PNG；PPTX 已保留。"
    result = pptx_path.with_suffix(".png")
    shutil.copyfile(images[0], result)
    return result, ""


def export_png(pptx_path: Path, render_dir: Path) -> tuple[Path | None, str]:
    system = platform.system()
    if system == "Darwin":
        image_path, warning = _export_png_with_keynote(pptx_path, render_dir)
        if image_path:
            return image_path, warning
        fallback_path, fallback_warning = _export_png_with_quicklook(pptx_path, render_dir)
        if fallback_path:
            return fallback_path, f"{warning} 已使用系统快速预览生成 PNG。"
        return None, f"{warning} {fallback_warning}".strip()
    if system == "Windows":
        return _export_png_with_powerpoint(pptx_path, render_dir)
    return None, "当前系统暂不支持自动转图，PPTX 已生成；请手动导出图片。"


def diagnostics_payload() -> dict[str, Any]:
    system = platform.system()
    template_count = len(_template_paths())
    checks: list[dict[str, Any]] = []
    warnings: list[str] = []
    actions: list[str] = []

    pptx_available = template_count > 0
    checks.append(
        {
            "key": "pptx",
            "label": "PPTX 生成",
            "ok": pptx_available,
            "level": "ok" if pptx_available else "error",
            "detail": f"已识别 {template_count} 份模板" if pptx_available else "未找到可用 PPTX 模板",
        }
    )

    png_available = False
    png_method = ""
    png_level = "error"
    png_label = "PNG 自动导出不可用"
    png_detail = "当前系统暂不支持自动转图。"

    if system == "Darwin":
        keynote_installed = Path("/Applications/Keynote.app").exists()
        keynote_script = KEYNOTE_EXPORT_SCRIPT.exists()
        quicklook_available = Path("/usr/bin/qlmanage").exists()
        checks.extend(
            [
                {
                    "key": "keynote",
                    "label": "Keynote",
                    "ok": keynote_installed,
                    "level": "ok" if keynote_installed else "warn",
                    "detail": "已安装" if keynote_installed else "未安装",
                },
                {
                    "key": "keynoteScript",
                    "label": "Keynote 导出脚本",
                    "ok": keynote_script,
                    "level": "ok" if keynote_script else "error",
                    "detail": "正常" if keynote_script else "缺失 scripts/export_keynote.applescript",
                },
                {
                    "key": "quicklook",
                    "label": "系统快速预览",
                    "ok": quicklook_available,
                    "level": "ok" if quicklook_available else "error",
                    "detail": "可作为 PNG 兜底导出" if quicklook_available else "未找到 qlmanage",
                },
            ]
        )
        if keynote_installed and keynote_script:
            png_available = True
            png_method = "Keynote"
            png_level = "ok"
            png_label = "PNG 自动导出可用"
            png_detail = "优先使用 Keynote 导出；若 Keynote 临时失败，会尝试系统快速预览兜底。"
            if quicklook_available:
                warnings.append("Keynote 需要 macOS 允许自动化控制；首次运行可能需要确认权限。")
        elif quicklook_available:
            png_available = True
            png_method = "Quick Look"
            png_level = "warn"
            png_label = "PNG 兜底导出可用"
            png_detail = "未满足 Keynote 自动导出条件，将使用系统快速预览生成 PNG。"
            warnings.append("Quick Look 兜底图可能和 Keynote/PowerPoint 渲染有细微差异。")
            if not keynote_installed:
                actions.append("如需更稳定的版式导出，建议安装 Keynote。")
            if not keynote_script:
                actions.append("检查 scripts/export_keynote.applescript 是否存在。")
        else:
            actions.append("安装 Keynote，或确认 /usr/bin/qlmanage 可用后再生成图片。")
    elif system == "Windows":
        powershell = _powershell_executable()
        script_ok = POWERPOINT_EXPORT_SCRIPT.exists()
        checks.extend(
            [
                {
                    "key": "powershell",
                    "label": "PowerShell",
                    "ok": bool(powershell),
                    "level": "ok" if powershell else "error",
                    "detail": powershell or "未找到",
                },
                {
                    "key": "powerpointScript",
                    "label": "PowerPoint 导出脚本",
                    "ok": script_ok,
                    "level": "ok" if script_ok else "error",
                    "detail": "正常" if script_ok else "缺失 scripts/export_powerpoint.ps1",
                },
            ]
        )
        if powershell and script_ok:
            png_available = True
            png_method = "PowerPoint"
            png_level = "warn"
            png_label = "PNG 自动导出待确认"
            png_detail = "PowerShell 和导出脚本可用；首次生成时会调用桌面版 PowerPoint。"
            warnings.append("请确保已安装桌面版 PowerPoint，并允许脚本自动化导出。")
        else:
            actions.append("安装桌面版 PowerPoint，并确认 PowerShell 与导出脚本可用。")
    else:
        checks.append(
            {
                "key": "system",
                "label": "系统支持",
                "ok": False,
                "level": "error",
                "detail": f"{system or '未知系统'} 暂不支持自动 PNG 导出",
            }
        )
        actions.append("可先下载 PPTX，再手动导出 PNG。")

    if not pptx_available:
        actions.append("将 PPTX 模板放到工具目录后点击“更新最新模板”。")

    return {
        "refreshedAt": datetime.now().isoformat(timespec="seconds"),
        "system": {
            "name": system,
            "release": platform.release(),
            "python": sys.version.split()[0],
        },
        "templates": {
            "count": template_count,
            "available": pptx_available,
        },
        "pptx": {
            "available": pptx_available,
            "level": "ok" if pptx_available else "error",
            "label": "PPTX 可生成" if pptx_available else "PPTX 暂不可生成",
            "detail": f"{template_count} 份模板可用" if pptx_available else "未找到可用模板",
        },
        "png": {
            "available": png_available,
            "method": png_method,
            "level": png_level,
            "label": png_label,
            "detail": png_detail,
        },
        "checks": checks,
        "warnings": warnings,
        "actions": actions,
    }


def _png_dimensions(image_bytes: bytes) -> tuple[int, int]:
    if len(image_bytes) < 24 or image_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("模板人物图片不是有效的 PNG 文件")
    width, height = struct.unpack(">II", image_bytes[16:24])
    if not width or not height:
        raise ValueError("模板人物图片尺寸不正确")
    return width, height


def _person_image_size(template_path: Path, media_path: str) -> tuple[int, int]:
    with zipfile.ZipFile(template_path) as archive:
        return _png_dimensions(archive.read(media_path))


def _shape_box(template_path: Path, shape_id: str) -> tuple[int, int, int, int]:
    with zipfile.ZipFile(template_path) as archive:
        slides = _slide_names(archive)
        root = ET.fromstring(archive.read(slides[0]))
    box = _shape_absolute_box_by_id(root, shape_id)
    if box is None:
        raise ValueError(f"模板图片框 {shape_id} 不存在")
    return box


def _upload_pixel_size(box: tuple[int, int, int, int]) -> tuple[int, int]:
    _, _, width, height = box
    if width >= height:
        return 1000, max(1, round(1000 * height / width))
    return max(1, round(1000 * width / height)), 1000


def _editor_fields_for_template(
    template_path: Path, values: dict[str, str]
) -> list[dict[str, Any]]:
    fields = fields_for_template(template_path)
    with zipfile.ZipFile(template_path) as archive:
        slide_width, slide_height = _slide_size(archive)
        slides = _slide_names(archive)
        if not slides or not slide_width or not slide_height:
            return []
        root = ET.fromstring(archive.read(slides[0]))

    editor_fields: list[dict[str, Any]] = []
    seen: set[tuple[str, int | None]] = set()
    for field in fields:
        box = _shape_absolute_box_by_id(root, field.shape_id)
        if box is None:
            continue
        identity = (field.shape_id, field.run_index)
        if identity in seen:
            continue
        seen.add(identity)
        x, y, width, height = box
        editor_fields.append(
            {
                "key": field.key,
                "label": field.label,
                "value": str(values.get(field.key, field.default)),
                "x": x / slide_width,
                "y": y / slide_height,
                "width": width / slide_width,
                "height": height / slide_height,
                "kind": field.kind,
            }
        )
    return editor_fields


def _sips_image_size(source_path: Path) -> tuple[int, int]:
    completed = subprocess.run(
        ["/usr/bin/sips", "-g", "pixelWidth", "-g", "pixelHeight", str(source_path)],
        check=False,
        capture_output=True,
        text=True,
        timeout=20,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()
        raise ValueError(f"人物照片读取失败：{detail or '请使用 PNG 或 JPG 图片'}")
    width_match = re.search(r"pixelWidth:\s*(\d+)", completed.stdout)
    height_match = re.search(r"pixelHeight:\s*(\d+)", completed.stdout)
    if not width_match or not height_match:
        raise ValueError("人物照片尺寸读取失败，请使用 PNG 或 JPG 图片")
    return int(width_match.group(1)), int(height_match.group(1))


def _normalize_image_on_macos(
    source_path: Path, png_path: Path, target_width: int, target_height: int
) -> None:
    source_width, source_height = _sips_image_size(source_path)
    resized_path = png_path.with_name("人物照片缩放.png")
    if source_width / source_height > target_width / target_height:
        resize_args = ["--resampleHeight", str(target_height)]
    else:
        resize_args = ["--resampleWidth", str(target_width)]
    commands = [
        [
            "/usr/bin/sips",
            "-s",
            "format",
            "png",
            *resize_args,
            str(source_path),
            "--out",
            str(resized_path),
        ],
        [
            "/usr/bin/sips",
            "--cropToHeightWidth",
            str(target_height),
            str(target_width),
            str(resized_path),
            "--out",
            str(png_path),
        ],
    ]
    try:
        for command in commands:
            completed = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
                timeout=45,
            )
            if completed.returncode != 0:
                detail = (completed.stderr or completed.stdout).strip()
                raise ValueError(
                    f"人物照片处理失败：{detail or '请使用 PNG 或 JPG 图片'}"
                )
    finally:
        resized_path.unlink(missing_ok=True)


def _convert_image_to_png(
    source_path: Path, png_path: Path, target_width: int, target_height: int
) -> None:
    system = platform.system()
    if system == "Darwin":
        _normalize_image_on_macos(source_path, png_path, target_width, target_height)
        return
    elif system == "Windows":
        powershell = _powershell_executable()
        if not WINDOWS_IMAGE_SCRIPT.exists() or not powershell:
            raise ValueError("Windows 图片转换组件不可用，请上传 PNG 图片")
        command = [
            powershell,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(WINDOWS_IMAGE_SCRIPT),
            "-SourcePath",
            str(source_path),
            "-TargetPath",
            str(png_path),
            "-TargetWidth",
            str(target_width),
            "-TargetHeight",
            str(target_height),
        ]
    else:
        raise ValueError("当前系统请上传 PNG 图片")
    completed = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        timeout=45,
    )
    if completed.returncode != 0 or not png_path.exists():
        detail = (completed.stderr or completed.stdout).strip()
        raise ValueError(f"人物照片转换失败：{detail or '请使用 PNG 或 JPG 图片'}")


def _prepare_person_image(
    photo_data_url: str,
    output_dir: Path,
    target_width: int,
    target_height: int,
    file_stem: str = "人物照片",
) -> bytes:
    match = re.fullmatch(r"data:image/([a-zA-Z0-9.+-]+);base64,(.+)", photo_data_url, re.S)
    if not match:
        raise ValueError("人物照片格式不正确，请重新上传")
    extension = re.sub(r"[^a-zA-Z0-9]+", "", match.group(1).lower()) or "image"
    try:
        image_bytes = base64.b64decode(match.group(2), validate=True)
    except binascii.Error as exc:
        raise ValueError("人物照片读取失败，请重新上传") from exc
    if not image_bytes:
        raise ValueError("人物照片为空，请重新上传")
    if len(image_bytes) > 12 * 1024 * 1024:
        raise ValueError("人物照片请控制在 12MB 以内")

    source_path = output_dir / f"{file_stem}-源文件.{extension}"
    png_path = output_dir / f"{file_stem}.png"
    source_path.write_bytes(image_bytes)
    _convert_image_to_png(source_path, png_path, target_width, target_height)
    if source_path != png_path:
        source_path.unlink(missing_ok=True)
    result = png_path.read_bytes()
    png_path.unlink(missing_ok=True)
    return result


def generate(payload: dict[str, Any]) -> dict[str, Any]:
    template_name = str(payload.get("template", ""))
    values = payload.get("values", {})
    if not isinstance(values, dict):
        raise ValueError("字段格式不正确")
    template_path = _template_path(template_name)
    GENERATED_DIR.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    stem = re.sub(r"[^\w\u4e00-\u9fff-]+", "-", template_path.stem).strip("-")
    output_dir = GENERATED_DIR / f"{stamp}-{stem}"
    suffix = 1
    while output_dir.exists():
        output_dir = GENERATED_DIR / f"{stamp}-{stem}-{suffix}"
        suffix += 1
    output_dir.mkdir()
    pptx_path = output_dir / f"{stem}-战报.pptx"
    media_replacements: dict[str, bytes] = {}
    gallery_media: list[tuple[str, str]] = []
    relink_images: list[tuple[str, str]] = []
    uploads = payload.get("uploads", {})
    if not isinstance(uploads, dict):
        raise ValueError("图片上传格式不正确")
    uploads = {str(key): str(value) for key, value in uploads.items() if value}
    text_styles = payload.get("textStyles", {})
    if not isinstance(text_styles, dict):
        raise ValueError("文案样式格式不正确")
    photo_data_url = payload.get("photo", "")
    if photo_data_url and "person" not in uploads:
        uploads["person"] = str(photo_data_url)
    person_photo = uploads.get("person", "")
    if person_photo:
        person_image_path = PERSON_IMAGE_BY_TEMPLATE.get(template_path.name)
        if not person_image_path:
            raise ValueError("该自动识别模板尚未配置人物照片替换位置")
        target_width, target_height = _person_image_size(template_path, person_image_path)
        media_replacements[person_image_path] = _prepare_person_image(
            person_photo, output_dir, target_width, target_height, "人物照片-1"
        )
    if template_path.name == "A200方案.pptx" and uploads.get("person2"):
        media_path = "ppt/media/upload-person-2.png"
        target_width, target_height = _person_image_size(
            template_path, PERSON_IMAGE_BY_TEMPLATE[template_path.name]
        )
        media_replacements[media_path] = _prepare_person_image(
            uploads["person2"], output_dir, target_width, target_height, "人物照片-2"
        )
        relink_images.append((A200_SECOND_PERSON_PICTURE_ID, media_path))
    if template_path.name == "Q1战报模版-横版.pptx":
        for index, shape_id in enumerate(HORIZONTAL_GALLERY_SHAPES, start=1):
            photo = uploads.get(f"gallery{index}", "")
            if not photo:
                continue
            media_path = f"ppt/media/upload-gallery-{index}.png"
            target_width, target_height = _upload_pixel_size(
                _shape_box(template_path, shape_id)
            )
            media_replacements[media_path] = _prepare_person_image(
                photo,
                output_dir,
                target_width,
                target_height,
                f"战报图片-{index}",
            )
            gallery_media.append((shape_id, media_path))
    clean_values = {str(k): str(v) for k, v in values.items()}
    render_pptx(
        template_path,
        clean_values,
        pptx_path,
        media_replacements,
        gallery_media,
        relink_images,
        text_styles,
    )
    image_path, warning = export_png(pptx_path, output_dir / "slide-render")
    editor_fields = _editor_fields_for_template(template_path, clean_values)
    record = _history_record(
        output_dir,
        template_path,
        clean_values,
        text_styles,
        pptx_path,
        image_path,
        warning,
        uploads,
        editor_fields,
    )
    _write_history_metadata(output_dir, record)
    response: dict[str, Any] = {
        "message": "战报已生成",
        "pptxUrl": _public_url(pptx_path),
        "imageUrl": _public_url(image_path) if image_path else "",
        "warning": warning,
        "outputDir": str(output_dir),
        "photoReplaced": bool(uploads),
        "uploadCount": len(uploads),
        "editorFields": editor_fields,
        "historyRecord": record,
    }
    return response


def _public_url(path: Path | None) -> str:
    if path is None:
        return ""
    relative = path.resolve().relative_to(ROOT)
    return "/" + urllib.parse.quote(relative.as_posix())


def _history_summary(values: dict[str, str]) -> str:
    preferred = ("name", "company", "amount", "headline", "product")
    picked = [
        str(values.get(key, "")).strip()
        for key in preferred
        if str(values.get(key, "")).strip()
    ][:2]
    if picked:
        return " · ".join(picked)
    for value in values.values():
        clean_value = str(value).strip()
        if clean_value:
            return clean_value[:42]
    return "未填写文案"


def _history_record(
    output_dir: Path,
    template_path: Path,
    values: dict[str, str],
    text_styles: dict[str, Any],
    pptx_path: Path,
    image_path: Path | None,
    warning: str,
    uploads: dict[str, str],
    editor_fields: list[dict[str, Any]],
    created_at: datetime | None = None,
) -> dict[str, Any]:
    created_at = created_at or datetime.now()
    return {
        "id": output_dir.name,
        "templateId": template_path.name,
        "templateName": template_path.stem,
        "imageUrl": _public_url(image_path) if image_path else "",
        "pptxUrl": _public_url(pptx_path),
        "warning": warning,
        "outputDir": str(output_dir),
        "photoReplaced": bool(uploads),
        "uploadCount": len(uploads),
        "editorFields": editor_fields,
        "values": values,
        "textStyles": text_styles,
        "summary": _history_summary(values),
        "createdAt": created_at.isoformat(timespec="seconds"),
    }


def _write_history_metadata(output_dir: Path, record: dict[str, Any]) -> None:
    metadata_path = output_dir / METADATA_FILENAME
    metadata_path.write_text(
        json.dumps(record, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _history_record_from_metadata(output_dir: Path) -> dict[str, Any] | None:
    metadata_path = output_dir / METADATA_FILENAME
    if not metadata_path.exists():
        return None
    try:
        data = json.loads(metadata_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    pptx_url = str(data.get("pptxUrl", ""))
    image_url = str(data.get("imageUrl", ""))
    if not pptx_url and not image_url:
        return None
    data["id"] = str(data.get("id") or output_dir.name)
    values = data.get("values", {})
    value_map = values if isinstance(values, dict) else {}
    data["summary"] = str(data.get("summary") or _history_summary(value_map))
    template_path = ROOT / str(data.get("templateId", ""))
    if template_path.exists():
        try:
            clean_values = {str(key): str(value) for key, value in value_map.items()}
            data["editorFields"] = _editor_fields_for_template(template_path, clean_values)
        except Exception:
            pass
    return data


def _history_record_from_output_dir(output_dir: Path) -> dict[str, Any] | None:
    pptx_files = sorted(output_dir.glob("*.pptx"))
    image_files = sorted(output_dir.glob("*.png"))
    if not pptx_files and not image_files:
        return None
    pptx_path = pptx_files[0] if pptx_files else None
    image_path = image_files[0] if image_files else None
    template_name = output_dir.name
    match = re.match(r"\d{8}-\d{6}-(.+)$", output_dir.name)
    if match:
        template_name = match.group(1)
    created_at = datetime.fromtimestamp(output_dir.stat().st_mtime)
    return {
        "id": output_dir.name,
        "templateId": f"{template_name}.pptx",
        "templateName": template_name,
        "imageUrl": _public_url(image_path) if image_path else "",
        "pptxUrl": _public_url(pptx_path) if pptx_path else "",
        "warning": "",
        "outputDir": str(output_dir),
        "photoReplaced": False,
        "uploadCount": 0,
        "editorFields": [],
        "values": {},
        "textStyles": {},
        "summary": template_name,
        "createdAt": created_at.isoformat(timespec="seconds"),
    }


def history_payload(limit: int = 20) -> dict[str, Any]:
    if not GENERATED_DIR.exists():
        return {"items": []}
    items: list[dict[str, Any]] = []
    output_dirs = sorted(
        (path for path in GENERATED_DIR.iterdir() if path.is_dir()),
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )
    for output_dir in output_dirs[: max(1, min(limit, 80))]:
        record = _history_record_from_metadata(output_dir) or _history_record_from_output_dir(output_dir)
        if record:
            items.append(record)
    return {"items": items}


def delete_history_record(record_id: str) -> dict[str, Any]:
    record_id = urllib.parse.unquote(str(record_id or "")).strip()
    if not record_id:
        raise ValueError("缺少历史记录 ID")
    candidate = (GENERATED_DIR / record_id).resolve()
    generated_root = GENERATED_DIR.resolve()
    if generated_root not in candidate.parents or not candidate.is_dir():
        raise ValueError("找不到要删除的历史记录")
    shutil.rmtree(candidate)
    return {"deleted": record_id, **history_payload(20)}


class AppHandler(SimpleHTTPRequestHandler):
    server_version = "BattleReportTool/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def _send_json(self, payload: Any, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/templates":
            try:
                self._send_json(templates_payload())
            except Exception as exc:
                self._send_json(
                    {"error": f"模板读取失败：{exc}"},
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )
            return
        if parsed.path == "/api/history":
            try:
                query = urllib.parse.parse_qs(parsed.query)
                limit = int(query.get("limit", ["20"])[0])
                self._send_json(history_payload(limit))
            except Exception as exc:
                self._send_json(
                    {"error": f"历史读取失败：{exc}"},
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )
            return
        if parsed.path == "/api/config":
            self._send_json(
                {
                    "shareMode": bool(getattr(self.server, "share_mode", False)),
                    "shareUrls": list(getattr(self.server, "share_urls", [])),
                }
            )
            return
        if parsed.path == "/api/diagnostics":
            try:
                self._send_json(diagnostics_payload())
            except Exception as exc:
                self._send_json(
                    {"error": f"环境诊断失败：{exc}"},
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )
            return
        if parsed.path == "/api/health":
            self._send_json({"ok": True})
            return
        if parsed.path == "/":
            self.path = "/static/index.html"
        super().do_GET()

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/api/generate":
            self._send_json({"error": "接口不存在"}, HTTPStatus.NOT_FOUND)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length > 110 * 1024 * 1024:
                raise ValueError("请求内容过大，单张图片请控制在 12MB 以内")
            payload = json.loads(self.rfile.read(length) or b"{}")
            with GENERATE_LOCK:
                self._send_json(generate(payload))
        except (ValueError, json.JSONDecodeError) as exc:
            self._send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self._send_json({"error": f"生成失败：{exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_DELETE(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        prefix = "/api/history/"
        if not parsed.path.startswith(prefix):
            self._send_json({"error": "接口不存在"}, HTTPStatus.NOT_FOUND)
            return
        try:
            record_id = parsed.path[len(prefix) :]
            self._send_json(delete_history_record(record_id))
        except ValueError as exc:
            self._send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self._send_json({"error": f"历史删除失败：{exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def translate_path(self, path: str) -> str:
        parsed_path = urllib.parse.urlparse(path).path
        relative = urllib.parse.unquote(parsed_path).lstrip("/")
        if relative.startswith("static/"):
            candidate = (ROOT / relative).resolve()
            if STATIC_DIR in candidate.parents:
                return str(candidate)
        if relative.startswith("generated/"):
            candidate = (ROOT / relative).resolve()
            if (
                GENERATED_DIR in candidate.parents
                and candidate.is_file()
                and candidate.suffix.lower() in {".png", ".pptx"}
                and "slide-render" not in candidate.parts
                and candidate.name != "人物照片.png"
            ):
                return str(candidate)
        return str(STATIC_DIR / "404")

    def list_directory(self, path: str) -> None:
        self.send_error(HTTPStatus.NOT_FOUND, "Directory listing is disabled")
        return None


def _make_server(preferred_port: int, host: str) -> tuple[ThreadingHTTPServer, int]:
    for port in range(preferred_port, preferred_port + 21):
        try:
            return ThreadingHTTPServer((host, port), AppHandler), port
        except OSError as exc:
            if exc.errno != errno.EADDRINUSE:
                raise
    raise OSError(f"端口 {preferred_port}-{preferred_port + 20} 均已被占用")


def _lan_addresses(port: int) -> list[str]:
    addresses: set[str] = set()
    try:
        for entry in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            addresses.add(entry[4][0])
    except OSError:
        pass
    try:
        probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        probe.connect(("8.8.8.8", 80))
        addresses.add(probe.getsockname()[0])
        probe.close()
    except OSError:
        pass
    return [
        f"http://{address}:{port}"
        for address in sorted(addresses)
        if address and not address.startswith("127.")
    ]


def serve(port: int, open_browser: bool, share: bool) -> None:
    GENERATED_DIR.mkdir(exist_ok=True)
    host = "0.0.0.0" if share else "127.0.0.1"
    server, actual_port = _make_server(port, host)
    url = f"http://127.0.0.1:{actual_port}"
    share_urls = _lan_addresses(actual_port) if share else []
    server.share_mode = share
    server.share_urls = share_urls
    print(f"战报生成工具已启动：{url}")
    if share:
        print("局域网共享模式已开启。其他同事可访问：")
        if share_urls:
            for share_url in share_urls:
                print(f"  {share_url}")
        else:
            print("  暂未识别到局域网地址，请检查网络连接。")
    print("按 Ctrl+C 停止服务。")
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
    finally:
        server.server_close()


def check_environment() -> int:
    print("战报生成工具环境检查")
    print(f"系统：{platform.system()} {platform.release()}")
    print(f"Python：{sys.version.split()[0]}")
    print(f"工具目录：{ROOT}")
    template_items = templates()
    print(f"PPTX 模板：{len(template_items)} 份")
    for item in template_items:
        print(f"  - {item['id']}（{item['orientation']}，{item['fieldCount']} 个字段）")
    if platform.system() == "Windows":
        print(f"Windows PowerShell：{_powershell_executable() or '未找到'}")
        print(f"PowerPoint 导出脚本：{'正常' if POWERPOINT_EXPORT_SCRIPT.exists() else '缺失'}")
        print("提示：桌面版 PowerPoint 会在首次生成图片时检查。")
    elif platform.system() == "Darwin":
        print(f"Keynote：{'已安装' if Path('/Applications/Keynote.app').exists() else '未安装'}")
    print("环境检查完成。")
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="从 PPTX 模板生成战报图片")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--share", action="store_true", help="允许局域网内其他电脑访问")
    parser.add_argument("--check", action="store_true", help="检查运行环境后退出")
    args = parser.parse_args()
    if args.check:
        raise SystemExit(check_environment())
    serve(args.port, not args.no_browser, args.share)
