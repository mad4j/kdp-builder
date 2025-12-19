from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class StyleDefinition:
    """Represents a style definition from YAML."""

    font_name: str = "Arial"
    font_size: int = 11
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None
    alignment: str = "left"
    space_before: int = 0
    space_after: int = 0

    @classmethod
    def from_dict(cls, style_data: Dict[str, Any]) -> "StyleDefinition":
        return cls(
            font_name=style_data.get("font_name", "Arial"),
            font_size=style_data.get("font_size", 11),
            bold=style_data.get("bold", False),
            italic=style_data.get("italic", False),
            underline=style_data.get("underline", False),
            color=style_data.get("color"),
            alignment=style_data.get("alignment", "left"),
            space_before=style_data.get("space_before", 0),
            space_after=style_data.get("space_after", 0),
        )


class LayoutDefinition:
    """Represents document layout from YAML."""

    MM_TO_INCHES = 1.0 / 25.4
    CM_TO_INCHES = 1.0 / 2.54

    def __init__(self, layout_data: Dict[str, Any]):
        self.unit = str(layout_data.get("unit", "inches")).lower()

        if self.unit not in ["inches", "mm", "cm"]:
            raise ValueError(
                f"Invalid unit '{self.unit}'. Must be 'inches', 'mm', or 'cm'."
            )

        self.page_width = self._convert_to_inches(layout_data.get("page_width", 8.5))
        self.page_height = self._convert_to_inches(layout_data.get("page_height", 11))
        self.margin_top = self._convert_to_inches(layout_data.get("margin_top", 1.0))
        self.margin_bottom = self._convert_to_inches(
            layout_data.get("margin_bottom", 1.0)
        )
        self.margin_left = self._convert_to_inches(layout_data.get("margin_left", 1.0))
        self.margin_right = self._convert_to_inches(
            layout_data.get("margin_right", 1.0)
        )
        self.header_text = layout_data.get("header_text")
        self.header_style = layout_data.get("header_style", "normal")
        self.footer_text = layout_data.get("footer_text")
        self.footer_style = layout_data.get("footer_style", "normal")

    def _convert_to_inches(self, value: float) -> float:
        if value < 0:
            raise ValueError(f"Dimension values must be positive, got {value}")
        if self.unit == "mm":
            return value * self.MM_TO_INCHES
        if self.unit == "cm":
            return value * self.CM_TO_INCHES
        return value
