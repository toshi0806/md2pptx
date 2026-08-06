#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""色名の解決（md2pptx．DESIGN.md §5.13）．

`accent1` のようなテーマ色名／`red` のような具体的な色名／`#ff0000` の16進を
1 か所で解く．python-pptx には依存しない純モジュールで，
戻り値は ``("theme", 名前)`` か ``("rgb", "RRGGBB")`` のどちらか——
テーマ色は**テーマ側で解決させたい**（差し替えに追従する）ので RGB へ潰さない．

**既定はテーマ色名**．テーマを差し替えても配色が破綻しないため．
具体的な色名と16進は「意味が色そのものに宿る」とき（危険＝赤／正常＝緑）に使う．

具体的な色名は **CSS Color Level 4 の named colors** をそのまま採る．
独自の色名表を作らないのは，作れば以後ずっと「なぜこの色が無いのか」に
答え続けることになるため．
"""
from __future__ import annotations

import re
from typing import Literal

# テーマ色名．描画側（render._theme_map）が MSO_THEME_COLOR へ対応付ける．
THEME_COLORS: frozenset[str] = frozenset({
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "tx1", "tx2", "bg1", "bg2",
})

# CSS Color Level 4 の named colors（148 色）．値は "RRGGBB"．
CSS_COLORS: dict[str, str] = {
    "aliceblue": "F0F8FF", "antiquewhite": "FAEBD7", "aqua": "00FFFF",
    "aquamarine": "7FFFD4", "azure": "F0FFFF", "beige": "F5F5DC",
    "bisque": "FFE4C4", "black": "000000", "blanchedalmond": "FFEBCD",
    "blue": "0000FF", "blueviolet": "8A2BE2", "brown": "A52A2A",
    "burlywood": "DEB887", "cadetblue": "5F9EA0", "chartreuse": "7FFF00",
    "chocolate": "D2691E", "coral": "FF7F50", "cornflowerblue": "6495ED",
    "cornsilk": "FFF8DC", "crimson": "DC143C", "cyan": "00FFFF",
    "darkblue": "00008B", "darkcyan": "008B8B", "darkgoldenrod": "B8860B",
    "darkgray": "A9A9A9", "darkgreen": "006400", "darkgrey": "A9A9A9",
    "darkkhaki": "BDB76B", "darkmagenta": "8B008B", "darkolivegreen": "556B2F",
    "darkorange": "FF8C00", "darkorchid": "9932CC", "darkred": "8B0000",
    "darksalmon": "E9967A", "darkseagreen": "8FBC8F", "darkslateblue": "483D8B",
    "darkslategray": "2F4F4F", "darkslategrey": "2F4F4F",
    "darkturquoise": "00CED1", "darkviolet": "9400D3", "deeppink": "FF1493",
    "deepskyblue": "00BFFF", "dimgray": "696969", "dimgrey": "696969",
    "dodgerblue": "1E90FF", "firebrick": "B22222", "floralwhite": "FFFAF0",
    "forestgreen": "228B22", "fuchsia": "FF00FF", "gainsboro": "DCDCDC",
    "ghostwhite": "F8F8FF", "gold": "FFD700", "goldenrod": "DAA520",
    "gray": "808080", "green": "008000", "greenyellow": "ADFF2F",
    "grey": "808080", "honeydew": "F0FFF0", "hotpink": "FF69B4",
    "indianred": "CD5C5C", "indigo": "4B0082", "ivory": "FFFFF0",
    "khaki": "F0E68C", "lavender": "E6E6FA", "lavenderblush": "FFF0F5",
    "lawngreen": "7CFC00", "lemonchiffon": "FFFACD", "lightblue": "ADD8E6",
    "lightcoral": "F08080", "lightcyan": "E0FFFF",
    "lightgoldenrodyellow": "FAFAD2", "lightgray": "D3D3D3",
    "lightgreen": "90EE90", "lightgrey": "D3D3D3", "lightpink": "FFB6C1",
    "lightsalmon": "FFA07A", "lightseagreen": "20B2AA",
    "lightskyblue": "87CEFA", "lightslategray": "778899",
    "lightslategrey": "778899", "lightsteelblue": "B0C4DE",
    "lightyellow": "FFFFE0", "lime": "00FF00", "limegreen": "32CD32",
    "linen": "FAF0E6", "magenta": "FF00FF", "maroon": "800000",
    "mediumaquamarine": "66CDAA", "mediumblue": "0000CD",
    "mediumorchid": "BA55D3", "mediumpurple": "9370DB",
    "mediumseagreen": "3CB371", "mediumslateblue": "7B68EE",
    "mediumspringgreen": "00FA9A", "mediumturquoise": "48D1CC",
    "mediumvioletred": "C71585", "midnightblue": "191970",
    "mintcream": "F5FFFA", "mistyrose": "FFE4E1", "moccasin": "FFE4B5",
    "navajowhite": "FFDEAD", "navy": "000080", "oldlace": "FDF5E6",
    "olive": "808000", "olivedrab": "6B8E23", "orange": "FFA500",
    "orangered": "FF4500", "orchid": "DA70D6", "palegoldenrod": "EEE8AA",
    "palegreen": "98FB98", "paleturquoise": "AFEEEE",
    "palevioletred": "DB7093", "papayawhip": "FFEFD5", "peachpuff": "FFDAB9",
    "peru": "CD853F", "pink": "FFC0CB", "plum": "DDA0DD",
    "powderblue": "B0E0E6", "purple": "800080", "rebeccapurple": "663399",
    "red": "FF0000", "rosybrown": "BC8F8F", "royalblue": "4169E1",
    "saddlebrown": "8B4513", "salmon": "FA8072", "sandybrown": "F4A460",
    "seagreen": "2E8B57", "seashell": "FFF5EE", "sienna": "A0522D",
    "silver": "C0C0C0", "skyblue": "87CEEB", "slateblue": "6A5ACD",
    "slategray": "708090", "slategrey": "708090", "snow": "FFFAFA",
    "springgreen": "00FF7F", "steelblue": "4682B4", "tan": "D2B48C",
    "teal": "008080", "thistle": "D8BFD8", "tomato": "FF6347",
    "turquoise": "40E0D0", "violet": "EE82EE", "wheat": "F5DEB3",
    "white": "FFFFFF", "whitesmoke": "F5F5F5", "yellow": "FFFF00",
    "yellowgreen": "9ACD32",
}

_RE_HEX = re.compile(r"^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")

ColorKind = Literal["theme", "rgb"]


def parse_color(value: str) -> tuple[ColorKind, str]:
    """色名を ``("theme", 名前)`` か ``("rgb", "RRGGBB")`` へ解決する．

    Args:
        value: テーマ色名（``accent1`` 等）／CSS の色名（``red`` 等）／
            16進（``#ff0000`` / ``#f00``）．前後の空白と大小文字は問わない．

    Returns:
        テーマ色なら ``("theme", 正規化した名前)``——**RGB へ潰さない**．
        テーマを差し替えたときに追従させたいので，解決はテーマ側に任せる．
        それ以外は ``("rgb", 大文字6桁)``．

    Raises:
        ValueError: どれにも当てはまらないとき．**黙って既定色にしない**——
            綴りを間違えた色は「効かない」ではなく「間違い」なので止める．
    """
    name = value.strip()
    lowered = name.lower()
    if lowered in THEME_COLORS:
        return "theme", lowered
    if lowered in CSS_COLORS:
        return "rgb", CSS_COLORS[lowered]
    m = _RE_HEX.match(name)
    if m:
        digits = m.group(1)
        if len(digits) == 3:            # #f00 → FF0000
            digits = "".join(c * 2 for c in digits)
        return "rgb", digits.upper()
    raise ValueError(
        f"unknown color {value!r} (use a theme color "
        f"{'/'.join(sorted(THEME_COLORS))}, a CSS color name like 'red', "
        f"or a hex value like '#ff0000')")
