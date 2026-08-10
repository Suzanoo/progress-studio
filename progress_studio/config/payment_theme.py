from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


CONFIG_PATH = Path(__file__).with_name("payment_lines.json")


@dataclass(frozen=True)
class PaymentLabelTheme:
    width_px: int = 145
    height_px: int = 26
    font_size: int = 12
    corner_radius_px: int = 6
    text_color: str = "FFFFFF"
    anchor_column_offset: int = -1
    anchor_row: int = 1
    collision_row_step: int = 1


@dataclass(frozen=True)
class PaymentLineTheme:
    colors: dict[str, str]
    line_style: str = "medium"
    endpoint_style: str = "thick"
    fallback_color: str = "7F7F7F"
    collision_max_offset: int = 3
    label: PaymentLabelTheme = PaymentLabelTheme()


DEFAULT_COLORS = {
    "P01": "C00000", "P02": "0070C0", "P03": "548235", "P04": "7030A0",
    "P05": "ED7D31", "P06": "00A6A6", "P07": "7F6000", "P08": "C0504D",
    "P09": "9BBB59", "P10": "4F81BD", "P11": "8064A2", "P12": "C4A000",
    "P13": "F06292", "P14": "008C95", "P15": "8C6D46",
}


def _hex(value: object, fallback: str) -> str:
    text = str(value or "").strip().lstrip("#").upper()
    if len(text) == 6 and all(ch in "0123456789ABCDEF" for ch in text):
        return text
    return fallback


def load_payment_line_theme(path: Path | None = None) -> PaymentLineTheme:
    source = Path(path) if path is not None else CONFIG_PATH
    data: dict = {}
    try:
        data = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        data = {}

    line = data.get("line") if isinstance(data.get("line"), dict) else {}
    label = data.get("label") if isinstance(data.get("label"), dict) else {}
    raw_colors = data.get("colors") if isinstance(data.get("colors"), dict) else {}

    colors = dict(DEFAULT_COLORS)
    for key, value in raw_colors.items():
        period_id = str(key).strip().upper()
        if period_id.startswith("P"):
            colors[period_id] = _hex(value, colors.get(period_id, "7F7F7F"))

    return PaymentLineTheme(
        colors=colors,
        line_style=str(line.get("style") or "medium"),
        endpoint_style=str(line.get("endpoint_style") or "thick"),
        fallback_color=_hex(line.get("fallback_color"), "7F7F7F"),
        collision_max_offset=max(int(line.get("collision_max_offset", 3)), 0),
        label=PaymentLabelTheme(
            width_px=max(int(label.get("width_px", 145)), 40),
            height_px=max(int(label.get("height_px", 26)), 14),
            font_size=max(int(label.get("font_size", 12)), 6),
            corner_radius_px=max(int(label.get("corner_radius_px", 6)), 0),
            text_color=_hex(label.get("text_color"), "FFFFFF"),
            anchor_column_offset=int(label.get("anchor_column_offset", -1)),
            anchor_row=max(int(label.get("anchor_row", 1)), 1),
            collision_row_step=max(int(label.get("collision_row_step", 1)), 0),
        ),
    )


# Compatibility constants for older callers/tests. New rendering loads the
# same config at renderer construction time.
_DEFAULT_THEME = load_payment_line_theme()
PAYMENT_LINE_COLORS = _DEFAULT_THEME.colors
PAYMENT_LINE_STYLE = _DEFAULT_THEME.line_style
