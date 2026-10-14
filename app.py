from flask import Flask, request, render_template_string
from itertools import permutations
import math
import re
import pandas as pd
from markupsafe import Markup
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024
HTML = """
<!doctype html>
<html lang="tr">
<head>
    <meta charset="utf-8">
    <meta
        name="viewport"
        content="width=device-width, initial-scale=1"
    >
    <title>Palet Koli Yerleşim Hesaplayıcı</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 1100px;
            margin: 35px auto;
            padding: 0 18px;
        }
        h1 {
            margin-bottom: 5px;
        }
        .note {
            color: #555;
            font-size: 0.92rem;
        }
        .section {
            margin-top: 25px;
            padding: 18px;
            border: 1px solid #bbb;
            border-radius: 10px;
        }
        .grid {
            display: grid;
            grid-template-columns: repeat(
                2,
                minmax(220px, 1fr)
            );
            gap: 14px;
        }
        label {
            display: block;
            font-weight: bold;
            margin-bottom: 5px;
        }
        input {
            width: 100%;
            padding: 9px;
            box-sizing: border-box;
        }
        button {
            margin-top: 18px;
            padding: 11px 18px;
            cursor: pointer;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
        }
        th,
        td {
            border: 1px solid #ccc;
            padding: 8px;
            text-align: center;
        }
        th {
            background: #f1f1f1;
        }
        .error {
            margin-top: 20px;
            padding: 14px;
            border: 1px solid #b00020;
            border-radius: 8px;
        }
        .drawing-card {
            margin-top: 25px;
            padding: 18px;
            border: 1px solid #bbb;
            border-radius: 8px;
            overflow-x: auto;
        }
        .technical-drawing {
            width: 100%;
            min-width: 650px;
            height: auto;
            background: white;
        }
        .pallet {
            fill: none;
            stroke: black;
            stroke-width: 3;
        }
        .box {
            fill: #f1f1f1;
            stroke: black;
            stroke-width: 1.5;
        }
        .box-number {
            font-size: 12px;
            text-anchor: middle;
            dominant-baseline: middle;
        }
        .dimension {
            stroke: #333;
            stroke-width: 1.5;
            marker-start: url(#arrow);
            marker-end: url(#arrow);
        }
        .dimension-text {
            font-size: 14px;
            text-anchor: middle;
        }
        @media (max-width: 650px) {
            .grid {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
<h1>Palet Koli Yerleşim Hesaplayıcı</h1>
<p class="note">
    Ölçüleri aynı birimde giriniz.
    Örneğin tümünü santimetre olarak giriniz.
</p>
<div class="section">
    <h2>Tekli Hesaplama</h2>
    <form
        method="post"
        enctype="multipart/form-data"
    >
        <input
            type="hidden"
            name="mode"
            value="single"
        >
        <div class="grid">
            <div>
                <label>Koli uzunluğu</label>
                <input
                    type="number"
                    step="0.01"
                    name="box_l"
                    value="{{ form.get('box_l', '') }}"
                >
            </div>
            <div>
                <label>Koli genişliği</label>
                <input
                    type="number"
                    step="0.01"
                    name="box_w"
                    value="{{ form.get('box_w', '') }}"
                >
            </div>
            <div>
                <label>Koli yüksekliği</label>
                <input
                    type="number"
                    step="0.01"
                    name="box_h"
                    value="{{ form.get('box_h', '') }}"
                >
            </div>
            <div>
                <label>Koli ağırlığı (kg)</label>
                <input
                    type="number"
                    step="0.01"
                    name="box_weight"
                    value="{{ form.get('box_weight', '') }}"
                >
            </div>
            <div>
                <label>Palet uzunluğu</label>
                <input
                    type="number"
                    step="0.01"
                    name="pallet_l"
                    value="{{ form.get('pallet_l', 120) }}"
                >
            </div>
            <div>
                <label>Palet genişliği</label>
                <input
                    type="number"
                    step="0.01"
                    name="pallet_w"
                    value="{{ form.get('pallet_w', 80) }}"
                >
            </div>
            <div>
                <label>İzin verilen toplam yükseklik</label>
                <input
                    type="number"
                    step="0.01"
                    name="max_h"
                    value="{{ form.get('max_h', 180) }}"
                >
            </div>
            <div>
                <label>Palet yüksekliği</label>
                <input
                    type="number"
                    step="0.01"
                    name="pallet_h"
                    value="{{ form.get('pallet_h', 15) }}"
                >
            </div>
            <div>
                <label>
                    Maksimum palet yükü
                    (kg, isteğe bağlı)
                </label>
                <input
                    type="number"
                    step="0.01"
                    name="max_weight"
                    value="{{ form.get('max_weight', '') }}"
                >
            </div>
<p class="note">
   Maksimum palet yükü 100x120lik palet için 1250 kg, 80x120lik palet için 1500 kg önerilir.
</p>
        </div>
        <button type="submit">
            Hesapla ve Teknik Çizimi Göster
        </button>
    </form>
</div>
<div class="section">
    <h2>Excel ile Toplu Hesaplama</h2>
    <p class="note">
        Excel sütun adları aşağıdaki şekilde olmalıdır:
    </p>
    <p class="note">
        Ürün Kodu,
        Koli Uzunluğu,
        Koli Genişliği,
        Koli Yüksekliği,
        Koli Ağırlığı,
        Palet Uzunluğu,
        Palet Genişliği,
        Maksimum Yükseklik,
        Palet Yüksekliği,
        Maksimum Palet Yükü
    </p>
    <form
        method="post"
        enctype="multipart/form-data"
    >
        <input
            type="hidden"
            name="mode"
            value="excel"
        >
        <label>Excel dosyası yükle</label>
        <input
            type="file"
            name="excel_file"
            accept=".xlsx"
        >
        <button type="submit">
            Excel'i Hesapla
        </button>
    </form>
</div>
{% if error %}
<div class="error">
    <strong>Hata:</strong>
    {{ error }}
</div>
{% endif %}
{% if result %}
<div class="section">
    <h2>Tekli Hesaplama Sonucu</h2>
    <table>
        <tr>
            <th>Koli tabanı</th>
            <th>Koli yüksekliği</th>
            <th>Boy yönünde</th>
            <th>En yönünde</th>
            <th>Bir katta</th>
            <th>Kat sayısı</th>
            <th>Toplam koli</th>
        </tr>
        <tr>
            <td>
                {{ result.base_l }}
                ×
                {{ result.base_w }}
            </td>
            <td>
                {{ result.height }}
            </td>
            <td>
                {{ result.boxes_along_length }}
            </td>
            <td>
                {{ result.boxes_along_width }}
            </td>
            <td>
                {{ result.per_layer }}
            </td>
            <td>
                {{ result.layers }}
            </td>
            <td>
                <strong>
                    {{ result.total }}
                </strong>
            </td>
        </tr>
    </table>
    <p>
        Palet taban kullanım oranı:
        <strong>
            %{{ result.area_utilization }}
        </strong>
    </p>
    <p>
        Hacim kullanım oranı:
        <strong>
            %{{ result.volume_utilization }}
        </strong>
    </p>
    {% if result.total_weight is not none %}
    <p>
        Toplam ürün ağırlığı:
        <strong>
            {{ result.total_weight }} kg
        </strong>
    </p>
    {% endif %}
    <div class="drawing-card">
        <h3>
            Teknik Çizim — 3B (İzometrik) Görünüş
        </h3>
        {{ iso_svg }}
        <p class="note">
            Mavi koli ölçü referansıdır:
            B<sub>L</sub> uzunluk,
            B<sub>W</sub> genişlik,
            B<sub>H</sub> yükseklik.
            Toplam {{ result.total }} koli,
            {{ result.layers }} kat.
        </p>
    </div>
    <div class="drawing-card">
        <h3>
            Teknik Çizim — Üstten Görünüş
        </h3>
        {% set scale = 4 %}
        {% set margin = 70 %}
        {% set pallet_draw_l =
            result.pallet_l * scale
        %}
        {% set pallet_draw_w =
            result.pallet_w * scale
        %}
        {% set box_draw_l =
            result.base_l * scale
        %}
        {% set box_draw_w =
            result.base_w * scale
        %}
        <svg
            class="technical-drawing"
            viewBox="
                0 0
                {{ pallet_draw_l + 150 }}
                {{ pallet_draw_w + 150 }}
            "
        >
            <defs>
                <marker
                    id="arrow"
                    markerWidth="8"
                    markerHeight="8"
                    refX="4"
                    refY="4"
                    orient="auto-start-reverse"
                >
                    <path
                        d="M 0 0 L 8 4 L 0 8 Z"
                    ></path>
                </marker>
            </defs>
            <rect
                x="{{ margin }}"
                y="{{ margin }}"
                width="{{ pallet_draw_l }}"
                height="{{ pallet_draw_w }}"
                class="pallet"
            ></rect>
            {% for row in
                range(result.boxes_along_width)
            %}
                {% for col in
                    range(result.boxes_along_length)
                %}
                    <rect
                        x="
                            {{
                                margin
                                + col * box_draw_l
                            }}
                        "
                        y="
                            {{
                                margin
                                + row * box_draw_w
                            }}
                        "
                        width="{{ box_draw_l }}"
                        height="{{ box_draw_w }}"
                        class="box"
                    ></rect>
                    <text
                        x="
                            {{
                                margin
                                + col * box_draw_l
                                + box_draw_l / 2
                            }}
                        "
                        y="
                            {{
                                margin
                                + row * box_draw_w
                                + box_draw_w / 2
                            }}
                        "
                        class="box-number"
                    >
                        {{
                            row
                            * result.boxes_along_length
                            + col
                            + 1
                        }}
                    </text>
                {% endfor %}
            {% endfor %}
            <line
                x1="{{ margin }}"
                y1="
                    {{
                        margin
                        + pallet_draw_w
                        + 30
                    }}
                "
                x2="
                    {{
                        margin
                        + pallet_draw_l
                    }}
                "
                y2="
                    {{
                        margin
                        + pallet_draw_w
                        + 30
                    }}
                "
                class="dimension"
            ></line>
            <text
                x="
                    {{
                        margin
                        + pallet_draw_l / 2
                    }}
                "
                y="
                    {{
                        margin
                        + pallet_draw_w
                        + 55
                    }}
                "
                class="dimension-text"
            >
                {{ result.pallet_l }} cm
            </text>
            <line
                x1="{{ margin - 30 }}"
                y1="{{ margin }}"
                x2="{{ margin - 30 }}"
                y2="
                    {{
                        margin
                        + pallet_draw_w
                    }}
                "
                class="dimension"
            ></line>
            <text
                x="{{ margin - 48 }}"
                y="
                    {{
                        margin
                        + pallet_draw_w / 2
                    }}
                "
                class="dimension-text"
                transform="
                    rotate(
                        -90
                        {{ margin - 48 }}
                        {{
                            margin
                            + pallet_draw_w / 2
                        }}
                    )
                "
            >
                {{ result.pallet_w }} cm
            </text>
        </svg>
        <p class="note">
            Koli tabanı:
            {{ result.base_l }}
            ×
            {{ result.base_w }}
            cm
        </p>
        <p class="note">
            Dizilim:
            {{ result.boxes_along_length }}
            ×
            {{ result.boxes_along_width }}
            =
            {{ result.per_layer }}
            koli/kat
        </p>
    </div>
</div>
{% endif %}
{% if bulk_table %}
<div class="section">
    <h2>
        Toplu Hesaplama Sonuçları
    </h2>
    {{ bulk_table | safe }}
</div>
{% endif %}
</body>
</html>
"""
def positive_float(
    form,
    key,
    optional=False,
    default=None
):
    raw = form.get(key, "").strip()
    if optional and raw == "":
        return default
    if raw == "":
        raise ValueError(
            f"{key} alanı boş bırakılamaz."
        )
    value = float(raw)
    if value <= 0:
        raise ValueError(
            f"{key} sıfırdan büyük olmalıdır."
        )
    return value
def safe_excel_float(
    row,
    column_name,
    optional=False,
    default=None
):
    value = row.get(column_name, None)
    if pd.isna(value) or value == "":
        if optional:
            return default
        raise ValueError(
            f"'{column_name}' sütununda boş değer var."
        )
    value = float(value)
    if value <= 0:
        raise ValueError(
            f"'{column_name}' değeri "
            "sıfırdan büyük olmalıdır."
        )
    return value
def calculate_best(
    box_dims,
    pallet_l,
    pallet_w,
    usable_h,
    pallet_h,
    box_weight=None,
    max_weight=None
):
    candidates = []
    for base_l, base_w, height in set(
        permutations(box_dims, 3)
    ):
        along_l = math.floor(
            pallet_l / base_l
        )
        along_w = math.floor(
            pallet_w / base_w
        )
        per_layer = along_l * along_w
        max_layers_by_height = math.floor(
            usable_h / height
        )
        if (
            per_layer <= 0
            or max_layers_by_height <= 0
        ):
            continue
        total = (
            per_layer
            * max_layers_by_height
        )
        if (
            box_weight is not None
            and max_weight is not None
        ):
            weight_limited_total = math.floor(
                max_weight / box_weight
            )
            total = min(
                total,
                weight_limited_total
            )
        if total <= 0:
            continue
        actual_layers = math.ceil(
            total / per_layer
        )
        area_utilization = (
            per_layer
            * base_l
            * base_w
            / (pallet_l * pallet_w)
            * 100
        )
        volume_utilization = (
            total
            * base_l
            * base_w
            * height
            / (
                pallet_l
                * pallet_w
                * usable_h
            )
            * 100
        )
        candidates.append({
            "base_l": round(
                base_l,
                2
            ),
            "base_w": round(
                base_w,
                2
            ),
            "height": round(
                height,
                2
            ),
            "boxes_along_length":
                along_l,
            "boxes_along_width":
                along_w,
            "per_layer":
                per_layer,
            "layers":
                actual_layers,
            "total":
                total,
            "pallet_l": round(
                pallet_l,
                2
            ),
            "pallet_w": round(
                pallet_w,
                2
            ),
            "pallet_h": round(
                pallet_h,
                2
            ),
            "max_h": round(
                usable_h + pallet_h,
                2
            ),
            "area_utilization": round(
                area_utilization,
                2
            ),
            "volume_utilization": round(
                volume_utilization,
                2
            ),
            "total_weight": (
                round(
                    total * box_weight,
                    2
                )
                if box_weight is not None
                else None
            )
        })
    if not candidates:
        return None
    return max(
        candidates,
        key=lambda item: (
            item["total"],
            item["area_utilization"],
            item["volume_utilization"]
        )
    )
ISO_COS = math.cos(math.radians(30))
ISO_SIN = 0.5
def iso_point(x, y, z):
    """x = pallet length, y = pallet width, z = height -> screen (sx, sy)."""
    return ((x + y) * ISO_COS, (x - y) * ISO_SIN - z)
def _pts(points):
    return " ".join(
        f"{px:.2f},{py:.2f}" for px, py in points
    )
def build_iso_svg(r):
    """Isometric (3D) pallet drawing like the reference picture."""
    L, W = r["pallet_l"], r["pallet_w"]
    ph, max_h = r["pallet_h"], r["max_h"]
    bl, bw, bh = r["base_l"], r["base_w"], r["height"]
    nx, per = r["boxes_along_length"], r["per_layer"]
    total, layers = r["total"], r["layers"]
    stack_top = ph + layers * bh
    top_z = max(stack_top, max_h)
    width = (L + W) * ISO_COS
    height = (L + W) * ISO_SIN + top_z
    min_y = -W * ISO_SIN - top_z
    ext = max(width, height)
    sw = ext * 0.0022
    fs = ext * 0.032
    d = ext * 0.05
    ah = ext * 0.022
    edge = "#3b2f1e"
    # ---- box grid -------------------------------------------------
    cells = {}
    for n in range(total):
        k, pos = divmod(n, per)
        j, i = divmod(pos, nx)
        cells[(i, j, k)] = n + 1
    # blue (highlighted) box = nearest visible one on top
    visible_top = [
        c for c in cells if (c[0], c[1], c[2] + 1) not in cells
    ]
    hi = max(visible_top, key=lambda c: (c[0] - c[1], c[2]))
    def face(points, fill):
        return (
            f'<polygon points="{_pts(points)}" fill="{fill}" '
            f'stroke="{edge}" stroke-width="{sw:.3f}" '
            f'stroke-linejoin="round"/>'
        )
    def block(x0, y0, z0, dx, dy, dz, colors, faces=("t", "l", "r")):
        x1, y1, z1 = x0 + dx, y0 + dy, z0 + dz
        P = iso_point
        out = []
        if "l" in faces:
            out.append(face([P(x0, y0, z0), P(x1, y0, z0),
                             P(x1, y0, z1), P(x0, y0, z1)],
                            colors[1]))
        if "r" in faces:
            out.append(face([P(x1, y0, z0), P(x1, y1, z0),
                             P(x1, y1, z1), P(x1, y0, z1)],
                            colors[2]))
        if "t" in faces:
            out.append(face([P(x0, y0, z1), P(x1, y0, z1),
                             P(x1, y1, z1), P(x0, y1, z1)],
                            colors[0]))
        return "".join(out)
    parts = []
    # ---- pallet ---------------------------------------------------
    wood = ("#c9a15a", "#b88a45", "#a37636")
    if ph > 0:
        parts.append(block(0, 0, 0, L, W, ph, wood))
        gap = "#4a3314"
        P = iso_point
        # openings on the long (left) side
        for a, b in ((0.08, 0.30), (0.38, 0.62), (0.70, 0.92)):
            parts.append(face([P(L*a, 0, ph*0.08), P(L*b, 0, ph*0.08),
                               P(L*b, 0, ph*0.70), P(L*a, 0, ph*0.70)],
                              gap))
        # openings on the short (right) side
        for a, b in ((0.10, 0.45), (0.55, 0.90)):
            parts.append(face([P(L, W*a, ph*0.08), P(L, W*b, ph*0.08),
                               P(L, W*b, ph*0.70), P(L, W*a, ph*0.70)],
                              gap))
    # ---- boxes (back to front, only visible faces) -----------------
    tan = ("#f2dcb3", "#e6c18a", "#d2a46a")
    blue = ("#8db4ee", "#6f9bdf", "#5b86cc")
    for (i, j, k) in sorted(
        cells, key=lambda c: (c[2] + c[0] - c[1], c[2])
    ):
        faces = []
        if (i, j, k + 1) not in cells:
            faces.append("t")
        if (i, j - 1, k) not in cells:
            faces.append("l")
        if (i + 1, j, k) not in cells:
            faces.append("r")
        if not faces and (i, j, k) != hi:
            continue
        colors = blue if (i, j, k) == hi else tan
        parts.append(
            block(i * bl, j * bw, ph + k * bh, bl, bw, bh,
                  colors, faces or ("t", "l", "r"))
        )
    # ---- dimension helpers ----------------------------------------
    def label(main, sub, rest):
        return (
            f'{main}<tspan dy="{fs*0.22:.2f}" font-size="{fs*0.68:.2f}">'
            f'{sub}</tspan><tspan dy="{-fs*0.22:.2f}">{rest}</tspan>'
        )
    def dim(p1, p2, off, text, anchor="middle", tdx=0, tdy=0,
            size=1.0):
        (x1, y1), (x2, y2) = p1, p2
        ox, oy = off
        a = (x1 + ox, y1 + oy)
        b = (x2 + ox, y2 + oy)
        mx, my = (a[0] + b[0]) / 2, (a[1] + b[1]) / 2
        # always use text-anchor=start (renders identically everywhere);
        # emulate "end"/"middle" with an estimated text width
        plain = re.sub(r"<[^>]+>", "", text)
        tw = len(plain) * fs * size * 0.55
        if anchor == "middle":
            tdx -= tw / 2
        elif anchor == "end":
            tdx -= tw
        anchor = "start"
        return (
            f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{a[0]:.2f}" '
            f'y2="{a[1]:.2f}" stroke="#555" stroke-width="{sw*0.6:.3f}"/>'
            f'<line x1="{x2:.2f}" y1="{y2:.2f}" x2="{b[0]:.2f}" '
            f'y2="{b[1]:.2f}" stroke="#555" stroke-width="{sw*0.6:.3f}"/>'
            f'<line x1="{a[0]:.2f}" y1="{a[1]:.2f}" x2="{b[0]:.2f}" '
            f'y2="{b[1]:.2f}" stroke="#111" stroke-width="{sw*1.2:.3f}" '
            f'marker-start="url(#iso-arrow)" marker-end="url(#iso-arrow)"/>'
            f'<text x="{mx+tdx:.2f}" y="{my+tdy:.2f}" '
            f'font-size="{fs*size:.2f}" text-anchor="{anchor}" '
            f'font-family="Arial, sans-serif" fill="#fff" '
            f'stroke="#fff" stroke-width="{sw*4:.3f}" '
            f'stroke-linejoin="round">{text}</text>'
            f'<text x="{mx+tdx:.2f}" y="{my+tdy:.2f}" '
            f'font-size="{fs*size:.2f}" text-anchor="{anchor}" '
            f'font-family="Arial, sans-serif" fill="#111">{text}</text>'
        )
    P = iso_point
    dims = []
    # pallet length / width / max height
    dims.append(dim(P(0, 0, 0), P(L, 0, 0), (-0.5*d*1.6, 0.866*d*1.6),
                    label("P", "L", f" = {L:g} cm"),
                    anchor="end", tdx=-fs*0.6, tdy=fs*0.9))
    dims.append(dim(P(L, 0, 0), P(L, W, 0), (0.5*d*1.6, 0.866*d*1.6),
                    label("P", "W", f" = {W:g} cm"),
                    anchor="start", tdx=fs*0.2, tdy=fs*0.5))
    dims.append(dim(P(L, W, 0), P(L, W, max_h), (d*1.6, 0),
                    label("H", "max", f" = {max_h:g} cm"),
                    anchor="start", tdx=fs*0.4, tdy=fs*0.35))
    # box length / width / height on the blue box
    i, j, k = hi
    x0, y0, z0 = i * bl, j * bw, ph + k * bh
    x1, y1, z1 = x0 + bl, y0 + bw, z0 + bh
    dims.append(dim(P(x0, y0, z0), P(x1, y0, z0), (-0.5*d*0.7, 0.866*d*0.7),
                    label("B", "L", f" = {bl:g}"),
                    anchor="end", tdx=-fs*0.5, tdy=fs*1.0, size=0.85))
    dims.append(dim(P(x1, y0, z0), P(x1, y1, z0), (0.5*d*0.7, 0.866*d*0.7),
                    label("B", "W", f" = {bw:g}"),
                    anchor="start", tdx=fs*0.2, tdy=fs*0.5, size=0.85))
    dims.append(dim(P(x1, y1, z0), P(x1, y1, z1), (d*0.7, 0),
                    label("B", "H", f" = {bh:g}"),
                    anchor="start", tdx=fs*0.3, tdy=fs*0.3, size=0.85))
    pad_l, pad_r = ext * 0.25, ext * 0.42
    pad_t, pad_b = ext * 0.04, ext * 0.14
    vb = (
        f"{-pad_l:.2f} {min_y - pad_t:.2f} "
        f"{width + pad_l + pad_r:.2f} {height + pad_t + pad_b:.2f}"
    )
    svg = (
        f'<svg class="technical-drawing" viewBox="{vb}" '
        f'xmlns="http://www.w3.org/2000/svg">'
        f'<defs><marker id="iso-arrow" viewBox="0 0 8 8" refX="8" '
        f'refY="4" markerUnits="userSpaceOnUse" markerWidth="{ah:.2f}" '
        f'markerHeight="{ah:.2f}" orient="auto-start-reverse">'
        f'<path d="M0 0 L8 4 L0 8 Z" fill="#111"/></marker></defs>'
        + "".join(parts)
        + "".join(dims)
        + "</svg>"
    )
    return Markup(svg)
@app.route(
    "/",
    methods=["GET", "POST"]
)
def index():
    result = None
    iso_svg = None
    bulk_table = None
    error = None
    form = (
        request.form
        if request.method == "POST"
        else {}
    )
    if request.method == "POST":
        mode = request.form.get(
            "mode",
            "single"
        )
        if mode == "excel":
            uploaded_file = request.files.get(
                "excel_file"
            )
            if (
                not uploaded_file
                or uploaded_file.filename == ""
            ):
                error = "Excel dosyası seçilmedi."
            else:
                try:
                    if not uploaded_file.filename.lower().endswith(".xlsx"):
                        raise ValueError("Yalnızca .xlsx dosyaları desteklenmektedir.")
                    df = pd.read_excel(
                        uploaded_file,
                        engine="openpyxl",
                        nrows=501
                    )
                    if len(df) > 500:
                        raise ValueError("En fazla 500 Excel veri satırı yükleyebilirsiniz.")
                    results = []
                    for (
                        row_number,
                        (_, row)
                    ) in enumerate(
                        df.iterrows(),
                        start=2
                    ):
                        try:
                            box_l = safe_excel_float(
                                row,
                                "Koli Uzunluğu"
                            )
                            box_w = safe_excel_float(
                                row,
                                "Koli Genişliği"
                            )
                            box_h = safe_excel_float(
                                row,
                                "Koli Yüksekliği"
                            )
                            box_weight = safe_excel_float(
                                row,
                                "Koli Ağırlığı",
                                optional=True,
                                default=None
                            )
                            pallet_l = safe_excel_float(
                                row,
                                "Palet Uzunluğu"
                            )
                            pallet_w = safe_excel_float(
                                row,
                                "Palet Genişliği"
                            )
                            max_h = safe_excel_float(
                                row,
                                "Maksimum Yükseklik"
                            )
                            pallet_h = safe_excel_float(
                                row,
                                "Palet Yüksekliği",
                                optional=True,
                                default=0
                            )
                            max_weight = safe_excel_float(
                                row,
                                "Maksimum Palet Yükü",
                                optional=True,
                                default=None
                            )
                            usable_h = (
                                max_h
                                - pallet_h
                            )
                            if usable_h <= 0:
                                raise ValueError(
                                    "Kullanılabilir "
                                    "yükseklik sıfırdan "
                                    "büyük olmalıdır."
                                )
                            row_result = calculate_best(
                                (
                                    box_l,
                                    box_w,
                                    box_h
                                ),
                                pallet_l,
                                pallet_w,
                                usable_h,
                                pallet_h,
                                box_weight,
                                max_weight
                            )
                            if row_result is None:
                                results.append({
                                    "Ürün Kodu":
                                        row.get(
                                            "Ürün Kodu",
                                            ""
                                        ),
                                    "Durum":
                                        "Yerleştirilemedi"
                                })
                            else:
                                results.append({
                                    "Ürün Kodu":
                                        row.get(
                                            "Ürün Kodu",
                                            ""
                                        ),
                                    "Koli Tabanı":
                                        (
                                            f"{row_result['base_l']} "
                                            f"× "
                                            f"{row_result['base_w']}"
                                        ),
                                    "Koli Yüksekliği":
                                        row_result[
                                            "height"
                                        ],
                                    "Boy Yönünde":
                                        row_result[
                                            "boxes_along_length"
                                        ],
                                    "En Yönünde":
                                        row_result[
                                            "boxes_along_width"
                                        ],
                                    "Bir Katta":
                                        row_result[
                                            "per_layer"
                                        ],
                                    "Kat Sayısı":
                                        row_result[
                                            "layers"
                                        ],
                                    "Toplam Koli":
                                        row_result[
                                            "total"
                                        ],
                                    "Alan Kullanımı (%)":
                                        row_result[
                                            "area_utilization"
                                        ],
                                    "Hacim Kullanımı (%)":
                                        row_result[
                                            "volume_utilization"
                                        ],
                                    "Toplam Ağırlık (kg)":
                                        row_result[
                                            "total_weight"
                                        ],
                                    "Durum":
                                        "Hesaplandı"
                                })
                        except Exception as row_exc:
                            results.append({
                                "Ürün Kodu":
                                    row.get(
                                        "Ürün Kodu",
                                        ""
                                    ),
                                "Durum":
                                    (
                                        f"{row_number}. satır "
                                        f"hatası: {row_exc}"
                                    )
                            })
                    result_df = pd.DataFrame(
                        results
                    )
                    bulk_table = result_df.to_html(
                        index=False,
                        border=0
                    )
                except Exception as exc:
                    error = (
                        f"Excel hesaplanamadı: "
                        f"{exc}"
                    )
        else:
            try:
                box_l = positive_float(
                    request.form,
                    "box_l"
                )
                box_w = positive_float(
                    request.form,
                    "box_w"
                )
                box_h = positive_float(
                    request.form,
                    "box_h"
                )
                pallet_l = positive_float(
                    request.form,
                    "pallet_l"
                )
                pallet_w = positive_float(
                    request.form,
                    "pallet_w"
                )
                max_h = positive_float(
                    request.form,
                    "max_h"
                )
                pallet_h = positive_float(
                    request.form,
                    "pallet_h",
                    optional=True,
                    default=0
                )
                box_weight = positive_float(
                    request.form,
                    "box_weight",
                    optional=True,
                    default=None
                )
                max_weight = positive_float(
                    request.form,
                    "max_weight",
                    optional=True,
                    default=None
                )
                usable_h = (
                    max_h
                    - pallet_h
                )
                if usable_h <= 0:
                    raise ValueError(
                        "Kullanılabilir yükseklik "
                        "sıfırdan büyük olmalıdır."
                    )
                result = calculate_best(
                    (
                        box_l,
                        box_w,
                        box_h
                    ),
                    pallet_l,
                    pallet_w,
                    usable_h,
                    pallet_h,
                    box_weight,
                    max_weight
                )
                if result is None:
                    error = (
                        "Bu ölçülerle palete "
                        "en az bir koli "
                        "yerleştirilemiyor."
                    )
            except ValueError as exc:
                error = str(exc)
            except Exception as exc:
                error = (
                    f"Program hatası: {exc}"
                )
    if result:
        iso_svg = build_iso_svg(result)
    return render_template_string(
        HTML,
        result=result,
        iso_svg=iso_svg,
        bulk_table=bulk_table,
        error=error,
        form=form
    )
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=8080,
        debug=False
    )
