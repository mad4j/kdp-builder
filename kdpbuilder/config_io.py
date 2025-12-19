from __future__ import annotations

from typing import Dict

import yaml

from .definitions import LayoutDefinition, StyleDefinition


def load_styles(yaml_path: str) -> Dict[str, StyleDefinition]:
    with open(yaml_path, "r", encoding="utf-8") as f:
        styles_data = yaml.safe_load(f)

    styles: Dict[str, StyleDefinition] = {}
    for style_name, style_data in (styles_data or {}).get("styles", {}).items():
        styles[style_name] = StyleDefinition.from_dict(style_data or {})

    return styles


def load_layout(yaml_path: str) -> LayoutDefinition:
    with open(yaml_path, "r", encoding="utf-8") as f:
        layout_data = yaml.safe_load(f)

    return LayoutDefinition((layout_data or {}).get("layout", {}))
