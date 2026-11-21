"""閲覧用スライドの配置。寸法は PowerPoint ポイントで、左上が原点。"""
from math import ceil, floor, log10
from pathlib import Path
from struct import unpack

from pptx_document import HEIGHT, WIDTH, Slide, cell_text, text_height


NAVY = "142D3C"
INK = "223743"
MUTED = "526777"
ACCENT = "087F77"
DARK_GREEN = "0F615C"
PALE = "EAF4F1"
PAPER = "F4F7F8"
RULE = "D9E4E8"
WHITE = "FFFFFF"
SERIES = (ACCENT, "4272A4", "A25830", "765797")

LEFT = 39
RIGHT = WIDTH - LEFT
CONTENT = RIGHT - LEFT
TOP = 111
BOTTOM = 495
BODY = 14
BODY_MIN = 12
TABLE_SIZE = 12


def put(slide, value, x, y, w, h, size=BODY, *, color=INK, bold=False,
        align="l", minimum=None, center=False, name=None):
    slide.text(value, x, y, w, h, size, color, bold, align,
               anchor="middle" if center else "top", name=name, minimum=minimum)


def rule(slide, y, x1=LEFT, x2=RIGHT, color=RULE, weight=0.7):
    slide.line(x1, y, x2, y, color, weight)


def heading(slide, label, x, y, width):
    put(slide, label, x, y, width, 24, 15, bold=True, color=DARK_GREEN, minimum=13)
    rule(slide, y + 26, x, x + width, weight=0.8)


def prose(slide, items, x, y, w, bottom):
    """段落の長さに応じて高さを配る。詰まったときだけ共通の書体サイズを下げる。"""
    for size in (BODY, 13, BODY_MIN):
        body_heights = [ceil(text_height(item["text"], w, size)) + 2 for item in items]
        title_heights = [ceil(text_height(item["heading"], w, size, True)) + 3
                         for item in items]
        if sum(body_heights) + sum(title_heights) + 11 * (len(items) - 1) <= bottom - y:
            break
    else:
        raise ValueError(f"説明が {BODY_MIN} pt で収まりません（{items[0]['heading']}）")
    for item, title_h, body_h in zip(items, title_heights, body_heights, strict=True):
        put(slide, item["heading"], x, y, w, title_h, size, bold=True)
        put(slide, item["text"], x, y + title_h, w, body_h, size, color=MUTED)
        y += title_h + body_h + 11


def row_pairs(slide, items, x, y, w, bottom, label_w=156, balanced=False):
    """短い項目名と説明の組を、折り返しに合わせて縦に配置する。"""
    gap = 14
    text_w = w - label_w - gap - 4
    for size in (BODY, 13, BODY_MIN):
        heights = [
            max(38, ceil(max(text_height(item["heading"], label_w, size, True),
                             text_height(item["text"], text_w, size))) + 14)
            for item in items
        ]
        if sum(heights) <= bottom - y:
            break
    else:
        raise ValueError(f"項目が {BODY_MIN} pt で収まりません（{items[0]['heading']}）")
    spare = (bottom - y - sum(heights)) / len(heights)
    heights = [h + (min(spare, 20) if balanced else 0) for h in heights]
    for item, h in zip(items, heights, strict=True):
        rule(slide, y, x, x + w)
        put(slide, item["heading"], x + 3, y + 6, label_w, h - 12, size,
            color=DARK_GREEN, bold=True, center=True)
        put(slide, item["text"], x + label_w + gap, y + 6, text_w, h - 12, size, center=True)
        y += h
    rule(slide, y, x, x + w)


def standard_slide(page, document, page_no):
    """表紙と章扉以外の共通見出し、資料名、出典を置く。"""
    slide = Slide(f"{page_no:02d} {page['title']}")
    put(slide, page.get("section", document.get("project", "")), LEFT, 20, 660, 16, 10,
        color=DARK_GREEN, bold=True)
    if page.get("continuation"):
        put(slide, f"表 {page['continuation']}", RIGHT - 135, 20, 135, 16, 10,
            color=DARK_GREEN, align="r")
    put(slide, page["title"], LEFT, 38, CONTENT, 62, 27, bold=True, color=NAVY, minimum=22,
        name="結論見出し")
    rule(slide, 102, color=ACCENT, weight=1.3)
    footer(slide, page, page_no)
    return slide


def footer(slide, page, number):
    rule(slide, 507)
    if page.get("source"):
        put(slide, page["source"], LEFT, 513, CONTENT - 43, 22, 9.5, color=MUTED,
            minimum=9, name="出典・注記")
    put(slide, f"{number:02d}", RIGHT - 36, 512, 36, 17, 10, color=MUTED, align="r")


def cover(slide, page, _root):
    slide.rect(0, 0, WIDTH, HEIGHT, PAPER)
    slide.rect(677, 0, WIDTH - 677, HEIGHT, NAVY)
    slide.rect(42, 66, 42, 4, ACCENT)
    put(slide, page.get("eyebrow", ""), 42, 80, 603, 37, 13, color=DARK_GREEN, bold=True)
    put(slide, page["title"], 42, 138, 588, 157, 41, color=NAVY, bold=True, minimum=32)
    if page.get("subtitle"):
        put(slide, page["subtitle"], 44, 309, 586, 81, 17, color=MUTED, minimum=15)
    if page.get("meta"):
        rule(slide, 454, 44, 627)
        put(slide, page["meta"], 44, 469, 583, 47, 12, color=MUTED, minimum=11)
    if page.get("topics"):
        put(slide, "本資料の範囲", 711, 134, 210, 26, 14, color="A5E2D8", bold=True)
    for i, label in enumerate(page.get("topics", [])):
        y = 178 + 58 * i
        rule(slide, y, 711, 923, color="4D6975")
        put(slide, label, 711, y + 11, 207, 44, 15, color=WHITE, minimum=13)


def contents(slide, page, _root):
    if page.get("intro"):
        put(slide, page["intro"], LEFT, TOP + 7, CONTENT, 42, 15, color=MUTED)
    y = TOP + (59 if page.get("intro") else 12)
    entries = page["items"]
    heights = [
        max(47, ceil(text_height(item["title"], 247, 15, True)) + 14,
            ceil(text_height(item.get("text", ""), 498, 13)) + 14)
        for item in entries
    ]
    if y + sum(heights) > BOTTOM:
        raise ValueError("目次の章名または説明が長すぎます")
    for i, (item, h) in enumerate(zip(entries, heights, strict=True), 1):
        rule(slide, y, color=RULE)
        put(slide, f"{i:02d}", LEFT, y + 6, 39, h - 12, 14, color=DARK_GREEN,
            bold=True, center=True)
        put(slide, item["title"], LEFT + 57, y + 6, 247, h - 12, 15,
            color=NAVY, bold=True, center=True)
        put(slide, item.get("text", ""), LEFT + 322, y + 6, 498, h - 12, 13,
            color=MUTED, center=True)
        put(slide, str(item["page"]), RIGHT - 39, y + 6, 39, h - 12, 13,
            color=NAVY, center=True, align="r")
        y += h
    rule(slide, y)


def section(slide, page, _root):
    slide.rect(0, 0, 562, HEIGHT, NAVY)
    slide.rect(0, 0, 7, HEIGHT, ACCENT)
    put(slide, page.get("eyebrow", "章の概要"), 44, 93, 478, 34, 13, color="A5E2D8", bold=True)
    put(slide, page["title"], 44, 140, 475, 137, 36, color=WHITE, bold=True, minimum=29)
    if page.get("intro"):
        put(slide, page["intro"], 44, 282, 472, 111, 16, color="D6E6EA", minimum=14)
    heading(slide, "この章の論点", 601, 111, 274)
    prose(slide, page["items"], 601, 155, 291, 480)


def summary(slide, page, _root):
    lead = page["lead"]
    lead_h = max(55, ceil(text_height(lead, CONTENT - 37, 17, True)) + 24)
    slide.rect(LEFT, TOP + 5, CONTENT, lead_h, PALE)
    slide.rect(LEFT, TOP + 5, 3, lead_h, ACCENT)
    put(slide, lead, LEFT + 17, TOP + 12, CONTENT - 37, lead_h - 14, 17,
        color=NAVY, bold=True, center=True)
    row_pairs(slide, page["items"], LEFT, TOP + lead_h + 23, CONTENT, BOTTOM,
              label_w=171, balanced=True)


def background(slide, page, _root):
    left_w = 524
    right_x = LEFT + left_w + 30
    heading(slide, "背景と目的", LEFT, TOP + 8, left_w)
    heading(slide, "この資料で扱う論点", right_x, TOP + 8, RIGHT - right_x)
    prose(slide, page["context"], LEFT, TOP + 49, left_w - 8, BOTTOM)
    prose(slide, page["questions"], right_x, TOP + 49, RIGHT - right_x, BOTTOM)


def table_widths(table):
    count = len(table["headers"])
    weights = table.get("widths", [1] * count)
    if len(weights) != count:
        raise ValueError("widths の個数は headers と同じにしてください")
    high = max(weights)
    normalized = [weight / high for weight in weights]
    widths = [CONTENT * weight / sum(normalized) for weight in normalized]
    if min(widths) < 48:
        raise ValueError("表の列幅が狭すぎます。widths の比率を調整してください")
    return widths


def table_row_height(widths, row, header=False, size=TABLE_SIZE):
    if len(row) != len(widths):
        raise ValueError("表の行の列数が headers と一致しません")
    estimate = max(
        text_height(cell_text(value), width - 16, size, header or j == 0)
        for j, (width, value) in enumerate(zip(widths, row, strict=True))
    )
    return max(30 if header else 28, ceil(estimate) + 10)


def table_area(page):
    intro = page.get("intro", "")
    note = page.get("note", "")
    start = TOP + 10
    if intro:
        start += ceil(text_height(intro, CONTENT, BODY)) + 16
    bottom = BOTTOM
    if note:
        bottom -= ceil(text_height(note, CONTENT - 29, 12)) + 29
    return start, bottom


def table_pages(page):
    """表の見出しと注記を確保したうえで、各行を切らずにページ分割する。"""
    widths = table_widths(page)
    y, end = table_area(page)
    header = table_row_height(widths, page["headers"], True)
    capacity = end - y - header
    groups = []
    rows = []
    used = 0
    for i, row in enumerate(page["rows"], 1):
        height = table_row_height(widths, row)
        if height > capacity:
            raise ValueError(f"表の {i} 行目が1ページに収まりません")
        if used + height > capacity:
            groups.append(rows)
            rows = []
            used = 0
        rows.append(row)
        used += height
    groups.append(rows)
    if len(groups) == 1:
        return [page]
    return [dict(page, rows=group, continuation=f"{i}/{len(groups)}")
            for i, group in enumerate(groups, 1)]


def render_table(slide, page, _root, *, emphasize_note=False):
    columns = len(page["headers"])
    widths = table_widths(page)
    rows = [page["headers"], *page["rows"]]
    heights = [table_row_height(widths, values, i == 0) for i, values in enumerate(rows)]
    y, end = table_area(page)
    if sum(heights) + y > end + 0.1:
        raise ValueError("表が収まりません。行数やセルの文章量を調整してください")
    if page.get("intro"):
        put(slide, page["intro"], LEFT, TOP + 4, CONTENT, y - TOP - 8, BODY, color=MUTED)
    highlight = page.get("highlight")
    aligns = page.get("aligns", ["l"] * columns)
    if len(aligns) != columns or (highlight is not None and highlight >= columns):
        raise ValueError("aligns または highlight が表の列数に合っていません")
    slide.table(LEFT, y, widths, heights, rows, NAVY, PALE, INK, highlight, TABLE_SIZE, aligns)
    if page.get("note"):
        note_y = y + sum(heights) + 12
        if emphasize_note:
            note_height = ceil(text_height(page["note"], CONTENT - 29, 12)) + 14
            slide.rect(LEFT, note_y, CONTENT, note_height, PALE)
            slide.rect(LEFT, note_y, 3, note_height, ACCENT)
            put(slide, page["note"], LEFT + 13, note_y + 5, CONTENT - 29, note_height - 10,
                12, color=DARK_GREEN, center=True)
        else:
            put(slide, page["note"], LEFT + 8, note_y, CONTENT - 17, BOTTOM - note_y, 12,
                color=MUTED)


def comparison(slide, page, root):
    render_table(slide, page, root, emphasize_note=True)


def method(slide, page, _root):
    # 左側の条件一覧と右側の検証手順は読み手が相互参照しやすいように並べる。
    left_w = 547
    right_x = LEFT + left_w + 30
    heading(slide, "測定条件", LEFT, TOP + 7, left_w)
    heading(slide, "実施・検証手順", right_x, TOP + 7, RIGHT - right_x)
    row_pairs(slide, page["conditions"], LEFT, TOP + 46, left_w, BOTTOM,
              label_w=121, balanced=True)
    prose(slide, page["steps"], right_x, TOP + 46, RIGHT - right_x, BOTTOM)


def tick_format(value):
    return f"{value:.5g}" if value else "0"


def axis_ticks(config, values):
    kind = config["kind"]
    lo, hi = min(values), max(values)
    if kind == "bar":
        lo, hi = min(0, lo), max(0, hi)
    if lo == hi:
        margin = max(1, abs(lo) * 0.15)
        lo, hi = lo - margin, hi + margin
    extent = hi - lo
    if extent <= 1e-280 or extent >= 1e280:
        raise ValueError("軸の値が極端です。千・万などに単位を変えてください")
    scale = 10 ** floor(log10(extent / 4))
    step = next(n * scale for n in (1, 2, 2.5, 5, 10) if n * scale >= extent / 4)
    lo = config.get("min", floor(lo / step) * step)
    hi = config.get("max", ceil(hi / step) * step)
    if hi <= lo or hi - lo >= 1e280 or (kind == "bar" and not lo <= 0 <= hi):
        raise ValueError("縦軸の範囲が不正です（棒グラフの範囲には0が必要です）")
    if min(values) < lo or max(values) > hi:
        raise ValueError("グラフ値が指定した縦軸の範囲を超えています")
    subdivisions = 4 if "min" in config or "max" in config else round((hi - lo) / step)
    return lo, hi, [lo + (hi - lo) * i / subdivisions for i in range(subdivisions + 1)]


def plot(slide, config, x, y, w, h):
    """棒・折れ線を PowerPoint の編集可能な線、図形、文字に変換する。"""
    series = config["series"]
    cats = config["categories"]
    if any(len(row["values"]) != len(cats) for row in series):
        raise ValueError(f"{config['title']} — categories と各系列の値の数が一致しません")
    all_values = [value for row in series for value in row["values"]]
    lo, hi, ticks = axis_ticks(config, all_values)
    put(slide, config["title"], x, y, w, 34, 15, bold=True, color=NAVY, minimum=13)
    put(slide, config.get("unit", ""), x, y + 37, 88, 13, 9.5, color=MUTED)
    px, pw = x + 52, w - 62
    py, ph = y + 56, h - 116

    def height_at(value):
        return py + ph * (hi - value) / (hi - lo)

    for tick in ticks:
        gy = height_at(tick)
        slide.line(px, gy, px + pw, gy, RULE, 0.6)
        put(slide, tick_format(tick), x, gy - 7, 44, 15, 9.5, color=MUTED,
            align="r", center=True)
    slide.line(px, py, px, py + ph, RULE, 0.8)
    slide.line(px, py + ph, px + pw, py + ph, RULE, 0.8)
    group_w = pw / len(cats)
    label_w = group_w - 2
    for i, cat in enumerate(cats):
        put(slide, cat, px + i * group_w + 1, py + ph + 5, label_w, 27,
            10, color=MUTED, align="ctr", minimum=9)
    for j, row in enumerate(series):
        color = SERIES[j]
        if config["kind"] == "line":
            points = [(px + (i + 0.5) * group_w, height_at(v))
                      for i, v in enumerate(row["values"])]
            for (x1, y1), (x2, y2) in zip(points, points[1:]):
                slide.line(x1, y1, x2, y2, color, 2.3)
            for cx, cy in points:
                slide.rect(cx - 2.6, cy - 2.6, 5.2, 5.2, color, geometry="ellipse")
        else:
            bar_w = min(42, group_w * 0.72 / len(series))
            zero = height_at(0)
            for i, value in enumerate(row["values"]):
                cx = px + (i + 0.5) * group_w + (j - (len(series) - 1) / 2) * bar_w
                vy = height_at(value)
                if value != 0:
                    slide.rect(cx - bar_w * 0.43, min(vy, zero), bar_w * 0.86, abs(vy - zero), color)
        legend_w = (w - 12) / len(series)
        lx = x + 7 + j * legend_w
        ly = y + h - 24
        slide.rect(lx, ly + 4, 12, 6, color)
        put(slide, row["label"], lx + 17, ly, legend_w - 19, 24, 10.5,
            color=MUTED, minimum=9.5)


def chart(slide, page, _root):
    charts = page["charts"]
    if len(charts) == 1:
        plot(slide, charts[0], LEFT, TOP + 5, 562, 368)
        right_x = LEFT + 582
        heading(slide, "読み取りと考察", right_x, TOP + 8, RIGHT - right_x)
        prose(slide, page["notes"], right_x, TOP + 47, RIGHT - right_x, BOTTOM)
        return
    half = (CONTENT - 28) / 2
    for i, chart_spec in enumerate(charts):
        plot(slide, chart_spec, LEFT + i * (half + 28), TOP + 5, half, 270)
    rule(slide, 396)
    notes = page["notes"]
    width = (CONTENT - 23 * (len(notes) - 1)) / len(notes)
    for i, item in enumerate(notes):
        prose(slide, [item], LEFT + i * (width + 23), 406, width, BOTTOM)


def image_size(data):
    """PNG と JPEG のヘッダーから元画像の寸法を得る。"""
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n" and data[12:16] == b"IHDR":
        return "png", *unpack(">II", data[16:24])
    if not data.startswith(b"\xff\xd8"):
        raise ValueError("画像は PNG または JPEG を指定してください")
    pos = 2
    while pos + 8 < len(data) and data[pos] == 255:
        marker = data[pos + 1]
        if marker == 255:
            pos += 1
            continue
        size = int.from_bytes(data[pos + 2:pos + 4], "big")
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            height, width = unpack(">HH", data[pos + 5:pos + 9])
            return "jpg", width, height
        if size < 2:
            break
        pos += size + 2
    raise ValueError("JPEG の画像寸法を読み込めません")


def annotated_image(slide, page, root):
    file = Path(page["image"])
    source = root / file
    try:
        data = source.read_bytes()
    except OSError as error:
        raise ValueError(f"画像を読めません（{file}）") from error
    ext, pixel_w, pixel_h = image_size(data)
    if min(pixel_w, pixel_h) <= 0:
        raise ValueError("画像の寸法が不正です")
    frame = (LEFT + 2, TOP + 10, 548, 316)
    ratio = min(frame[2] / pixel_w, frame[3] / pixel_h)
    iw, ih = pixel_w * ratio, pixel_h * ratio
    ix = frame[0] + (frame[2] - iw) / 2
    iy = frame[1] + (frame[3] - ih) / 2
    slide.rect(*frame, PAPER, RULE)
    slide.picture(data, ext, ix, iy, iw, ih, page["alt"])
    notes = page["annotations"]
    for i, item in enumerate(notes, 1):
        cx, cy = ix + iw * item["x"], iy + ih * item["y"]
        slide.rect(cx - 11, cy - 11, 22, 22, ACCENT, WHITE, geometry="ellipse")
        put(slide, f"{i}", cx - 9, cy - 10, 18, 20, 10.5, color=WHITE, align="ctr",
            bold=True, center=True)
    if page.get("caption"):
        put(slide, page["caption"], LEFT, TOP + 337, 550, 42, 11, color=MUTED)
    right_x = LEFT + 577
    heading(slide, "図の読み方", right_x, TOP + 9, RIGHT - right_x)
    items = [dict(heading=f"{i}　{note['heading']}", text=note["text"])
             for i, note in enumerate(notes, 1)]
    prose(slide, items, right_x, TOP + 49, RIGHT - right_x, BOTTOM)


def case(slide, page, _root):
    left_w = 328
    right_x = LEFT + left_w + 29
    heading(slide, "事例の概要", LEFT, TOP + 8, left_w)
    heading(slide, "取組・結果・適用条件", right_x, TOP + 8, RIGHT - right_x)
    row_pairs(slide, page["profile"], LEFT, TOP + 49, left_w, BOTTOM,
              label_w=68, balanced=True)
    prose(slide, page["findings"], right_x, TOP + 49, RIGHT - right_x, BOTTOM)


def structure(slide, page, _root):
    columns = page["columns"]
    nodes = page["nodes"]
    row_count = max(node["row"] for node in nodes) + 1
    row_height = min(88, (310 - 20 * (row_count - 1)) / row_count)
    gap = 30
    w = (CONTENT - gap * (len(columns) - 1)) / len(columns)
    cells = {}
    places = {}
    for node in nodes:
        col, row, key = node["col"], node["row"], node["id"]
        if col >= len(columns) or key in places or (col, row) in cells:
            raise ValueError(f"図の配置またはノードIDが重複・範囲外です（{key}）")
        x, y = LEFT + (w + gap) * col, TOP + 52 + (row_height + 22) * row
        places[key] = (col, row, x, y)
        cells[col, row] = key

    for i, label in enumerate(columns):
        heading(slide, label, LEFT + (w + gap) * i, TOP + 8, w)
    for edge in page["links"]:
        if edge["from"] not in places or edge["to"] not in places:
            raise ValueError("構造図の矢印に未定義のノードがあります")
        c1, r1, x1, y1 = places[edge["from"]]
        c2, r2, x2, y2 = places[edge["to"]]
        if c1 == c2 and abs(r1 - r2) == 1:
            cy1, cy2 = (y1 + row_height, y2) if r1 < r2 else (y1, y2 + row_height)
            slide.line(x1 + w / 2, cy1, x2 + w / 2, cy2, ACCENT, 1.5, arrow=True)
        elif abs(c1 - c2) == 1:
            start, end = (x1 + w, x2) if c1 < c2 else (x1, x2 + w)
            mid = (start + end) / 2
            cy1, cy2 = y1 + row_height / 2, y2 + row_height / 2
            if r1 == r2:
                slide.line(start, cy1, end, cy2, ACCENT, 1.5, arrow=True)
            else:
                slide.line(start, cy1, mid, cy1, ACCENT, 1.5)
                slide.line(mid, cy1, mid, cy2, ACCENT, 1.5)
                slide.line(mid, cy2, end, cy2, ACCENT, 1.5, arrow=True)
        else:
            raise ValueError("矢印は隣の列、または同じ列の隣の行を結んでください")
    for node in nodes:
        _, _, x, y = places[node["id"]]
        slide.rect(x, y, w, row_height, PALE, RULE)
        put(slide, node["label"], x + 10, y + 4, w - 20, row_height - 8, 14,
            color=NAVY, bold=True, minimum=12, align="ctr", center=True)
    if page.get("note"):
        put(slide, page["note"], LEFT, BOTTOM - 23, CONTENT, 23, 11.5, color=MUTED)


def roadmap(slide, page, _root):
    periods = page["periods"]
    tasks = page["tasks"]
    left_w = 208
    cell_w = (CONTENT - left_w) / len(periods)
    y = TOP + 12
    row_h = min(57, (306 - 30) / len(tasks))
    put(slide, "作業・担当", LEFT + 8, y, left_w - 16, 24, 13, color=DARK_GREEN, bold=True)
    for i, label in enumerate(periods):
        put(slide, label, LEFT + left_w + i * cell_w, y, cell_w, 24, 13,
            color=DARK_GREEN, bold=True, align="ctr")
    y += 33
    for i, task in enumerate(tasks):
        start, end = task["start"], task["end"]
        if not 0 <= start <= end < len(periods):
            raise ValueError(f"期間の番号が範囲外です（{task['task']}）")
        ry = y + i * row_h
        if i % 2 == 0:
            slide.rect(LEFT, ry, CONTENT, row_h, PAPER)
        rule(slide, ry)
        put(slide, task["task"], LEFT + 8, ry + 4, left_w - 17, 21, 13, bold=True, minimum=11.5)
        put(slide, task.get("owner", ""), LEFT + 8, ry + 25, left_w - 17, row_h - 27,
            10.5, color=MUTED)
        for j in range(len(periods)):
            gx = LEFT + left_w + j * cell_w
            slide.line(gx, ry, gx, ry + row_h, RULE, 0.6)
        bx = LEFT + left_w + start * cell_w + 4
        bw = (end - start + 1) * cell_w - 8
        slide.rect(bx, ry + 8, bw, row_h - 16, ACCENT)
        put(slide, task.get("outcome", ""), bx + 7, ry + 11, bw - 14, row_h - 22, 11,
            color=WHITE, minimum=10, center=True)
    rule(slide, y + len(tasks) * row_h)
    if page.get("note"):
        put(slide, page["note"], LEFT, BOTTOM - 38, CONTENT, 38, 12, color=MUTED)


def issues(slide, page, _root):
    columns = [(LEFT + 11, 308), (LEFT + 351, 391), (RIGHT - 104, 93)]
    y = TOP + 9
    if page.get("intro"):
        h = ceil(text_height(page["intro"], CONTENT, BODY)) + 3
        put(slide, page["intro"], LEFT, y, CONTENT, h, color=MUTED)
        y += h + 15
    slide.rect(LEFT, y, CONTENT, 29, NAVY)
    for (x, w), label in zip(columns, ("課題・影響", "対応策", "担当"), strict=True):
        put(slide, label, x, y + 3, w, 22, 13, bold=True, color=WHITE, center=True)
    y += 29
    bottom = BOTTOM
    if page.get("note"):
        bottom -= ceil(text_height(page["note"], CONTENT - 15, 12)) + 23
    rows = page["items"]
    for size in (13, 12):
        fits = []
        for row in rows:
            label_height = ceil(text_height(row["issue"], columns[0][1], size, True)) + 2
            issue_height = label_height + ceil(text_height(row["impact"], columns[0][1], size)) + 4
            action_height = ceil(text_height(row["action"], columns[1][1], size))
            owner_height = ceil(text_height(row["owner"], columns[2][1], size, True))
            fits.append((label_height, max(45, issue_height + 12, action_height + 12, owner_height + 12)))
        if sum(height for _, height in fits) <= bottom - y:
            break
    else:
        raise ValueError(f"課題と対応策が {BODY_MIN} pt で収まりません")
    spacing = min(8, (bottom - y - sum(height for _, height in fits)) / len(rows))
    for i, (row, (label_h, h)) in enumerate(zip(rows, fits, strict=True)):
        h += spacing
        if i % 2 == 0:
            slide.rect(LEFT, y, CONTENT, h, PAPER)
        slide.rect(LEFT, y, 2, h, ACCENT)
        put(slide, row["issue"], columns[0][0], y + 6, columns[0][1], label_h, size,
            bold=True, color=NAVY)
        put(slide, row["impact"], columns[0][0], y + 6 + label_h, columns[0][1],
            h - 12 - label_h, size, color=MUTED)
        for key, (x, w) in zip(("action", "owner"), columns[1:], strict=True):
            put(slide, row[key], x, y + 6, w, h - 12, size, center=True,
                color=DARK_GREEN if key == "owner" else INK, bold=key == "owner")
        y += h
        rule(slide, y)
    if page.get("note"):
        put(slide, page["note"], LEFT + 7, y + 12, CONTENT - 15, BOTTOM - y - 12, 12, color=MUTED)


RENDERERS = {
    "cover": cover,
    "contents": contents,
    "section": section,
    "summary": summary,
    "background": background,
    "method": method,
    "chart": chart,
    "comparison": comparison,
    "data_table": render_table,
    "annotated_image": annotated_image,
    "case": case,
    "structure": structure,
    "roadmap": roadmap,
    "issues": issues,
}
