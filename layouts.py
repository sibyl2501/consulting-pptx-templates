"""閲覧用スライドの配置。座標は 960 × 540 の基準キャンバス上で、左上が原点。"""

# ========================================
# ライブラリの読み込み
# ========================================
from math import ceil, floor, log10
from pathlib import Path
from struct import unpack

from pptx_document import DEFAULT_FONT, DEFAULT_PAGE_SIZE, HEIGHT, WIDTH, Slide, cell_text, text_height


# ========================================
# 色の指定
# ========================================
# 配色は # を除いた 6 桁の RGB 値。NAVY は主見出し、INK は本文、MUTED は補足に使う。
# ACCENT は強調、DARK_GREEN は小見出し、PALE・PAPER は背景、RULE は罫線用。
# SERIES はグラフの系列順の配色で、凡例とデータの色をそろえる。
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

# ========================================
# 余白と本文領域の指定
# ========================================
# x は右向き、y は下向き。LEFT〜RIGHT と本文の開始位置〜BOTTOM が主な本文領域。
# 本文領域の外に共通見出しと出典・ページ番号を置く。
LEFT = 39
RIGHT = WIDTH - LEFT
CONTENT = RIGHT - LEFT
BOTTOM = 495

# 見出しの座標と文字サイズを保ち、1行で収まるときは罫線を見出しの近くに置く。
TITLE_Y = 38
TITLE_HEIGHT = 62
TITLE_SIZE = 27
TITLE_MIN = 22
TITLE_RULE_LIMIT = 102
BODY_GAP = 9

# ========================================
# 文字サイズの指定
# ========================================
# BODY は本文の基準、BODY_MIN は縮小の下限、TABLE_SIZE は表の行高を測る際の文字サイズ。
BODY = 12
BODY_MIN = 8
TABLE_SIZE = 8


# ========================================
# 文字・罫線・段落の共通処理
# ========================================
# テキストボックスを置く。(x, y) は左上、(w, h) は幅と高さ。
# align は横方向（l・ctr・r）、center は縦方向の中央寄せ。minimum を指定すると、そのサイズまで縮小できる。
def put(slide, value, x, y, w, h, size=BODY, *, color=INK, bold=False,
        align="l", minimum=None, center=False, name=None):
    slide.text(value, x, y, w, h, size, color, bold, align,
               anchor="middle" if center else "top", name=name, minimum=minimum)


# 水平の罫線を置く。x1〜x2 を部分的な範囲にすれば、段組みの中でも使える。
def rule(slide, y, x1=LEFT, x2=RIGHT, color=RULE, weight=0.7):
    slide.line(x1, y, x2, y, color, weight)


# 小見出しと直下の罫線をひと組にし、列の左端と幅をそろえる。
def heading(slide, label, x, y, width):
    put(slide, label, x, y, width, 24, 15, bold=True, color=DARK_GREEN, minimum=13)
    rule(slide, y + 26, x, x + width, weight=0.8)


def prose(slide, items, x, y, w, bottom):
    """段落の長さに応じて高さを配る。詰まったときだけ共通の書体サイズを下げる。"""
    for size in (BODY, 13, BODY_MIN):
        # 見出しと本文の折り返し高さを別々に測る。段落間に取る 11 も含め、文字サイズの候補が収まるか試す。
        body_heights = [ceil(text_height(item["text"], w, size)) + 2 for item in items]
        title_heights = [ceil(text_height(item["heading"], w, size, True)) + 3
                         for item in items]
        if sum(body_heights) + sum(title_heights) + 11 * (len(items) - 1) <= bottom - y:
            break
    else:
        raise ValueError(f"説明が {BODY_MIN} pt で収まりません（{items[0]['heading']}）")
    # どの大きさも収まらなければ for の else でエラーになる。収まれば、見出しと本文を上下に置く。
    for item, title_h, body_h in zip(items, title_heights, body_heights, strict=True):
        put(slide, item["heading"], x, y, w, title_h, size, bold=True)
        put(slide, item["text"], x, y + title_h, w, body_h, size, color=MUTED)
        y += title_h + body_h + 11


def row_pairs(slide, items, x, y, w, bottom, label_w=156, balanced=False):
    """短い項目名と説明の組を、折り返しに合わせて縦に配置する。"""
    # 先にラベル幅と列間の隙間を確保し、残りを説明欄とする。
    # 行高は左右で折り返し後に高い方を基準に決める。
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
    # balanced が有効なら余った高さを均等に配分する。1 行の追加分には 20 の上限を設ける。
    spare = (bottom - y - sum(heights)) / len(heights)
    heights = [h + (min(spare, 20) if balanced else 0) for h in heights]
    for item, h in zip(items, heights, strict=True):
        rule(slide, y, x, x + w)
        put(slide, item["heading"], x + 3, y + 6, label_w, h - 12, size,
            color=DARK_GREEN, bold=True, center=True)
        put(slide, item["text"], x + label_w + gap, y + 6, text_w, h - 12, size, center=True)
        y += h
    rule(slide, y, x, x + w)


# ========================================
# 共通見出しとフッター
# ========================================
def title_rule_y(title):
    # 標準サイズで折り返す場合は、従来の位置に罫線を置いて見出しの高さを確保する。
    height = text_height(title, CONTENT, TITLE_SIZE, True)
    one_line_height = text_height("", CONTENT, TITLE_SIZE, True)
    if height > one_line_height:
        return TITLE_RULE_LIMIT
    return min(TITLE_RULE_LIMIT, TITLE_Y + ceil(height) + 7)


def body_top(page):
    # 改ページを計算する表も同じ位置から始める。
    return title_rule_y(page["title"]) + BODY_GAP


def standard_slide(page, document, page_no, page_size=DEFAULT_PAGE_SIZE, *, font=DEFAULT_FONT):
    """表紙と章扉以外の共通見出し、資料名、出典を置く。"""
    # スライド名は通し番号とタイトル。見出し上の小さなラベルは section があれば章名、なければ資料名。
    slide = Slide(f"{page_no:02d} {page['title']}", page_size, font=font)
    put(slide, page.get("section", document.get("project", "")), LEFT, 20, 660, 16, 10,
        color=DARK_GREEN, bold=True)
    # 詳細表が分割された場合は、右上に「表 2/3」のような続き番号を置く。
    if page.get("continuation"):
        put(slide, f"表 {page['continuation']}", RIGHT - 135, 20, 135, 16, 10,
            color=DARK_GREEN, align="r")
    put(slide, page["title"], LEFT, TITLE_Y, CONTENT, TITLE_HEIGHT, TITLE_SIZE,
        bold=True, color=NAVY, minimum=TITLE_MIN, name="結論見出し")
    rule(slide, title_rule_y(page["title"]), color=ACCENT, weight=1.3)
    footer(slide, page, page_no)
    return slide


# 出典は任意。ページ番号に必要な幅を残し、出典欄は残りの幅を使う。
def footer(slide, page, number):
    rule(slide, 507)
    if page.get("source"):
        put(slide, page["source"], LEFT, 513, CONTENT - 43, 22, 9.5, color=MUTED,
            minimum=9, name="出典・注記")
    put(slide, f"{number:02d}", RIGHT - 36, 512, 36, 17, 10, color=MUTED, align="r")


# ========================================
# 表紙・目次・章扉
# ========================================
# 表紙は右の濃色帯に topics、左に表題と副題などを置く。
# _root は画像フォルダを受ける共通引数で、画像を使わない形式では参照しない。
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
    # 話題の区切りは各項目を同じ高さでそろえる。
    for i, label in enumerate(page.get("topics", [])):
        y = 178 + 58 * i
        rule(slide, y, 711, 923, color="4D6975")
        put(slide, label, 711, y + 11, 207, 44, 15, color=WHITE, minimum=13)


# 目次。前置きの有無で表の開始位置を変え、章名と説明の折り返しを測って行高を決める。
def contents(slide, page, _root):
    top = body_top(page)
    if page.get("intro"):
        put(slide, page["intro"], LEFT, top + 7, CONTENT, 42, 15, color=MUTED)
    y = top + (59 if page.get("intro") else 12)
    entries = page["items"]
    heights = [
        max(47, ceil(text_height(item["title"], 247, 15, True)) + 14,
            ceil(text_height(item.get("text", ""), 498, 13)) + 14)
        for item in entries
    ]
    if y + sum(heights) > BOTTOM:
        raise ValueError("目次の章名または説明が長すぎます")
    # 左端の通し番号と右端の参照ページは別。参照先は入力の item["page"] を使う。
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


# 章扉の左に章名と導入、右にこの章で扱う論点（items）を置く。
def section(slide, page, _root):
    slide.rect(0, 0, 562, HEIGHT, NAVY)
    slide.rect(0, 0, 7, HEIGHT, ACCENT)
    put(slide, page.get("eyebrow", "章の概要"), 44, 93, 478, 34, 13, color="A5E2D8", bold=True)
    put(slide, page["title"], 44, 140, 475, 137, 36, color=WHITE, bold=True, minimum=29)
    if page.get("intro"):
        put(slide, page["intro"], 44, 282, 472, 111, 16, color="D6E6EA", minimum=14)
    heading(slide, "この章の論点", 601, 111, 274)
    prose(slide, page["items"], 601, 155, 291, 480)


# ========================================
# 要約と背景のレイアウト
# ========================================
# 要約。lead の折り返しに合わせて結論欄を高くし、その直下に根拠や補足を並べる。
def summary(slide, page, _root):
    top = body_top(page)
    lead = page["lead"]
    lead_h = max(55, ceil(text_height(lead, CONTENT - 37, 17, True)) + 24)
    slide.rect(LEFT, top + 5, CONTENT, lead_h, PALE)
    slide.rect(LEFT, top + 5, 3, lead_h, ACCENT)
    put(slide, lead, LEFT + 17, top + 12, CONTENT - 37, lead_h - 14, 17,
        color=NAVY, bold=True, center=True)
    row_pairs(slide, page["items"], LEFT, top + lead_h + 23, CONTENT, BOTTOM,
              label_w=171, balanced=True)


# 背景と論点。左列を広く取り、context と questions を別々の段落として並べる。
def background(slide, page, _root):
    top = body_top(page)
    left_w = 524
    right_x = LEFT + left_w + 30
    heading(slide, "背景と目的", LEFT, top + 8, left_w)
    heading(slide, "この資料で扱う論点", right_x, top + 8, RIGHT - right_x)
    prose(slide, page["context"], LEFT, top + 49, left_w - 8, BOTTOM)
    prose(slide, page["questions"], right_x, top + 49, RIGHT - right_x, BOTTOM)


# ========================================
# 表の列幅・行高・改ページと描画
# ========================================
# widths は列幅の比。省略時は等幅で、合計が本文幅の CONTENT になるよう配分する。
def table_widths(table):
    count = len(table["headers"])
    weights = table.get("widths", [1] * count)
    if len(weights) != count:
        raise ValueError("widths の個数は headers と同じにしてください")
    # 先に最大値で割り、大きな重み同士を直接加算して扱える範囲を超えるのを防ぐ。
    high = max(weights)
    normalized = [weight / high for weight in weights]
    widths = [CONTENT * weight / sum(normalized) for weight in normalized]
    if min(widths) < 48:
        raise ValueError("表の列幅が狭すぎます。widths の比率を調整してください")
    return widths


# 最も高さが要るセルに行高を合わせる。見出し行と先頭列は太字として見積もる。
# 文字を測る幅からはセル左右の余白を引き、戻り値には上下の余白も足す。
def table_row_height(widths, row, header=False, size=TABLE_SIZE):
    if len(row) != len(widths):
        raise ValueError("表の行の列数が headers と一致しません")
    estimate = max(
        text_height(cell_text(value), width - 16, size, header or j == 0)
        for j, (width, value) in enumerate(zip(widths, row, strict=True))
    )
    return max(30 if header else 28, ceil(estimate) + 10)


# 前置きの intro と注記の note に場所を取り、実際に表を置ける y の開始・終了位置を返す。
def table_area(page):
    top = body_top(page)
    intro = page.get("intro", "")
    note = page.get("note", "")
    start = top + 10
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
    # 見出し行を分割後のページでも繰り返すため、本文行に使える高さから見出し分を除く。
    capacity = end - y - header
    groups = []
    rows = []
    used = 0
    for i, row in enumerate(page["rows"], 1):
        height = table_row_height(widths, row)
        if height > capacity:
            raise ValueError(f"表の {i} 行目が1ページに収まりません")
        # 追加するとあふれるときは、今のページを確定し、その行を次のページの先頭に入れる。
        if used + height > capacity:
            groups.append(rows)
            rows = []
            used = 0
        rows.append(row)
        used += height
    groups.append(rows)
    if len(groups) == 1:
        return [page]
    # 元の辞書をコピーし、rows だけをページごとに差し替える。continuation は出力時に付ける印。
    return [dict(page, rows=group, continuation=f"{i}/{len(groups)}")
            for i, group in enumerate(groups, 1)]


# 比較表と詳細表で共用する表描画。詳細表を分割する処理は table_pages が先に行う。
def render_table(slide, page, _root, *, emphasize_note=False):
    top = body_top(page)
    columns = len(page["headers"])
    widths = table_widths(page)
    rows = [page["headers"], *page["rows"]]
    heights = [table_row_height(widths, values, i == 0) for i, values in enumerate(rows)]
    y, end = table_area(page)
    # 寸法計算の丸め誤差を考慮して、領域との比較には小さな許容幅を設ける。
    if sum(heights) + y > end + 0.1:
        raise ValueError("表が収まりません。行数やセルの文章量を調整してください")
    if page.get("intro"):
        put(slide, page["intro"], LEFT, top + 4, CONTENT, y - top - 8, BODY, color=MUTED)
    # 強調列は 0 始まり。列ごとの左右・中央寄せの個数も実際の列数と照合する。
    highlight = page.get("highlight")
    aligns = page.get("aligns", ["l"] * columns)
    if len(aligns) != columns or (highlight is not None and highlight >= columns):
        raise ValueError("aligns または highlight が表の列数に合っていません")
    # 寸法も含めて Slide に渡す。出力は PowerPoint 上でセルごとに編集できる表になる。
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


# 共通の表描画を呼び、比較表の注記だけは帯で強調する。
def comparison(slide, page, root):
    render_table(slide, page, root, emphasize_note=True)


# ========================================
# 調査・実験方法のレイアウト
# ========================================
def method(slide, page, _root):
    top = body_top(page)
    # 左側の条件一覧と右側の検証手順は読み手が相互参照しやすいように並べる。
    left_w = 547
    right_x = LEFT + left_w + 30
    heading(slide, "測定条件", LEFT, top + 7, left_w)
    heading(slide, "実施・検証手順", right_x, top + 7, RIGHT - right_x)
    row_pairs(slide, page["conditions"], LEFT, top + 46, left_w, BOTTOM,
              label_w=121, balanced=True)
    prose(slide, page["steps"], right_x, top + 46, RIGHT - right_x, BOTTOM)


# 目盛りの値は有効数字 5 桁に整え、0 は「0」と表示する。

# ========================================
# グラフの目盛りと描画
# ========================================
def tick_format(value):
    return f"{value:.5g}" if value else "0"


# 全系列を対象に縦軸の上下限と目盛りの位置を求める。
def axis_ticks(config, values):
    kind = config["kind"]
    lo, hi = min(values), max(values)
    # 棒の長さを 0 からの差として読めるよう、軸に 0 を含める。全値が同じ場合は上下に余白を取る。
    if kind == "bar":
        lo, hi = min(0, lo), max(0, hi)
    if lo == hi:
        margin = max(1, abs(lo) * 0.15)
        lo, hi = lo - margin, hi + margin
    extent = hi - lo
    if extent <= 1e-280 or extent >= 1e280:
        raise ValueError("軸の値が極端です。千・万などに単位を変えてください")
    # 約 4 区間を目安に、10 の累乗に 1・2・2.5・5・10 を掛けた幅から刻みを選ぶ。
    # 自動設定の端は刻み幅の倍数に向かって外側に丸める。
    scale = 10 ** floor(log10(extent / 4))
    step = next(n * scale for n in (1, 2, 2.5, 5, 10) if n * scale >= extent / 4)
    lo = config.get("min", floor(lo / step) * step)
    hi = config.get("max", ceil(hi / step) * step)
    if hi <= lo or hi - lo >= 1e280 or (kind == "bar" and not lo <= 0 <= hi):
        raise ValueError("縦軸の範囲が不正です（棒グラフの範囲には0が必要です）")
    if min(values) < lo or max(values) > hi:
        raise ValueError("グラフ値が指定した縦軸の範囲を超えています")
    # 軸端が指定されていれば範囲を 4 等分し、自動の場合は選んだ刻み幅を使う。
    subdivisions = 4 if "min" in config or "max" in config else round((hi - lo) / step)
    return lo, hi, [lo + (hi - lo) * i / subdivisions for i in range(subdivisions + 1)]


def plot(slide, config, x, y, w, h):
    """棒・折れ線を PowerPoint の編集可能な線、図形、文字に変換する。"""
    series = config["series"]
    # 横軸の各項目と系列の値は同じ順で対応する。件数が違う場合はグラフ名を付けて止める。
    cats = config["categories"]
    if any(len(row["values"]) != len(cats) for row in series):
        raise ValueError(f"{config['title']} — categories と各系列の値の数が一致しません")
    all_values = [value for row in series for value in row["values"]]
    lo, hi, ticks = axis_ticks(config, all_values)
    put(slide, config["title"], x, y, w, 34, 15, bold=True, color=NAVY, minimum=13)
    put(slide, config.get("unit", ""), x, y + 37, 88, 13, 9.5, color=MUTED)
    # 枠内の左に目盛り、上にグラフ名と単位、下にラベルと凡例の空間を取る。
    # (px, py) は値を描く範囲の左上、(pw, ph) は幅と高さ。
    px, pw = x + 52, w - 62
    py, ph = y + 56, h - 116

    # y は下向きに増えるので、最大値 hi を上端、最小値 lo を下端に対応させる。
    def height_at(value):
        return py + ph * (hi - value) / (hi - lo)

    for tick in ticks:
        gy = height_at(tick)
        slide.line(px, gy, px + pw, gy, RULE, 0.6)
        put(slide, tick_format(tick), x, gy - 7, 44, 15, 9.5, color=MUTED,
            align="r", center=True)
    slide.line(px, py, px, py + ph, RULE, 0.8)
    slide.line(px, py + ph, px + pw, py + ph, RULE, 0.8)
    # 横軸を項目数で等分し、項目ラベルや折れ線の点は区画の中央を基準にする。
    group_w = pw / len(cats)
    label_w = group_w - 2
    for i, cat in enumerate(cats):
        put(slide, cat, px + i * group_w + 1, py + ph + 5, label_w, 27,
            10, color=MUTED, align="ctr", minimum=9)
    for j, row in enumerate(series):
        color = SERIES[j]
        # 折れ線では隣り合う点をつなぎ、点の位置を示す丸を重ねる。棒は系列ごとに左右にずらす。
        if config["kind"] == "line":
            points = [(px + (i + 0.5) * group_w, height_at(v))
                      for i, v in enumerate(row["values"])]
            for (x1, y1), (x2, y2) in zip(points, points[1:]):
                slide.line(x1, y1, x2, y2, color, 2.3)
            for cx, cy in points:
                slide.rect(cx - 2.6, cy - 2.6, 5.2, 5.2, color, geometry="ellipse")
        else:
            bar_w = min(42, group_w * 0.72 / len(series))
            # 正の棒はゼロから上、負の棒は下へ伸びる。上側の座標と長さを min と絶対値で計算する。
            zero = height_at(0)
            for i, value in enumerate(row["values"]):
                cx = px + (i + 0.5) * group_w + (j - (len(series) - 1) / 2) * bar_w
                vy = height_at(value)
                if value != 0:
                    slide.rect(cx - bar_w * 0.43, min(vy, zero), bar_w * 0.86, abs(vy - zero), color)
        # グラフの下を系列数で分け、系列色の矩形とラベルを順に置く。
        legend_w = (w - 12) / len(series)
        lx = x + 7 + j * legend_w
        ly = y + h - 24
        slide.rect(lx, ly + 4, 12, 6, color)
        put(slide, row["label"], lx + 17, ly, legend_w - 19, 24, 10.5,
            color=MUTED, minimum=9.5)


# グラフが 1 点なら右に考察、2 点ならグラフを左右に並べて下に考察を置く。
def chart(slide, page, _root):
    top = body_top(page)
    charts = page["charts"]
    if len(charts) == 1:
        plot(slide, charts[0], LEFT, top + 5, 562, BOTTOM - top - 16)
        right_x = LEFT + 582
        heading(slide, "読み取りと考察", right_x, top + 8, RIGHT - right_x)
        prose(slide, page["notes"], right_x, top + 47, RIGHT - right_x, BOTTOM)
        return
    half = (CONTENT - 28) / 2
    notes_rule_y = 396
    for i, chart_spec in enumerate(charts):
        plot(slide, chart_spec, LEFT + i * (half + 28), top + 5, half, notes_rule_y - top - 15)
    rule(slide, notes_rule_y)
    notes = page["notes"]
    # グラフが 2 点ある場合、考察の件数で残りの幅を等分し、列間に 23 の間隔を設ける。
    width = (CONTENT - 23 * (len(notes) - 1)) / len(notes)
    for i, item in enumerate(notes):
        prose(slide, [item], LEFT + i * (width + 23), notes_rule_y + 10, width, BOTTOM)


# ========================================
# 画像の読み込みと注釈
# ========================================
def image_size(data):
    """PNG と JPEG のヘッダーから元画像の寸法を得る。"""
    # PNG の IHDR から幅と高さを、上位バイトから並ぶ 4 バイト整数として読む。
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n" and data[12:16] == b"IHDR":
        return "png", *unpack(">II", data[16:24])
    if not data.startswith(b"\xff\xd8"):
        raise ValueError("画像は PNG または JPEG を指定してください")
    # JPEG は先頭マーカーの後から項目長を手掛かりに SOF（画像寸法を持つ項目）まで進む。
    # SOF 内では高さ、幅の順なので、PNG と同様の戻り値（形式、幅、高さ）に整える。
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


# 相対パスは root（入力 JSON の場所）を基準にする。
# 画像を埋め込み、その上の番号と右側の説明を別要素として描く。
def annotated_image(slide, page, root):
    top = body_top(page)
    file = Path(page["image"])
    source = root / file
    try:
        data = source.read_bytes()
    except OSError as error:
        raise ValueError(f"画像を読めません（{file}）") from error
    ext, pixel_w, pixel_h = image_size(data)
    if min(pixel_w, pixel_h) <= 0:
        raise ValueError("画像の寸法が不正です")
    # 表示枠の中に元画像の縦横比で収め、余った空間を左右または上下に均等に取る。
    image_bottom = 437
    frame = (LEFT + 2, top + 10, 548, image_bottom - top - 10)
    iw, ih = slide.fit_picture(pixel_w, pixel_h, frame[2], frame[3])
    ix = frame[0] + (frame[2] - iw) / 2
    iy = frame[1] + (frame[3] - ih) / 2
    slide.rect(*frame, PAPER, RULE)
    slide.picture(data, ext, ix, iy, iw, ih, page["alt"])
    # 注釈の x と y は 0〜1 の比率。外側の枠ではなく、実際の画像の端を基準にする。
    notes = page["annotations"]
    for i, item in enumerate(notes, 1):
        cx, cy = ix + iw * item["x"], iy + ih * item["y"]
        slide.rect(cx - 11, cy - 11, 22, 22, ACCENT, WHITE, geometry="ellipse")
        put(slide, f"{i}", cx - 9, cy - 10, 18, 20, 10.5, color=WHITE, align="ctr",
            bold=True, center=True)
    if page.get("caption"):
        put(slide, page["caption"], LEFT, image_bottom + 11, 550, 42, 11, color=MUTED)
    right_x = LEFT + 577
    heading(slide, "図の読み方", right_x, top + 9, RIGHT - right_x)
    items = [dict(heading=f"{i}　{note['heading']}", text=note["text"])
             for i, note in enumerate(notes, 1)]
    prose(slide, items, right_x, top + 49, RIGHT - right_x, BOTTOM)


# 事例。左の profile は項目名と値の組、右の findings は見出し付きの説明として並べる。

# ========================================
# 事例紹介のレイアウト
# ========================================
def case(slide, page, _root):
    top = body_top(page)
    left_w = 328
    right_x = LEFT + left_w + 29
    heading(slide, "事例の概要", LEFT, top + 8, left_w)
    heading(slide, "取組・結果・適用条件", right_x, top + 8, RIGHT - right_x)
    row_pairs(slide, page["profile"], LEFT, top + 49, left_w, BOTTOM,
              label_w=68, balanced=True)
    prose(slide, page["findings"], right_x, top + 49, RIGHT - right_x, BOTTOM)


# 関係図。行・列は 0 始まり。ノードの ID を使って links の出発点と到着点を特定する。

# ========================================
# 構造図のレイアウト
# ========================================
def structure(slide, page, _root):
    top = body_top(page)
    columns = page["columns"]
    nodes = page["nodes"]
    row_count = max(node["row"] for node in nodes) + 1
    row_height = min(88, (310 - 20 * (row_count - 1)) / row_count)
    gap = 30
    w = (CONTENT - gap * (len(columns) - 1)) / len(columns)
    # cells は格子位置の重複検査、places はノード ID から配置座標を引くために使う。
    cells = {}
    places = {}
    for node in nodes:
        col, row, key = node["col"], node["row"], node["id"]
        if col >= len(columns) or key in places or (col, row) in cells:
            raise ValueError(f"図の配置またはノードIDが重複・範囲外です（{key}）")
        x, y = LEFT + (w + gap) * col, top + 52 + (row_height + 22) * row
        places[key] = (col, row, x, y)
        cells[col, row] = key

    for i, label in enumerate(columns):
        heading(slide, label, LEFT + (w + gap) * i, top + 8, w)
    # 線はカードの外辺から出入りする。結べるのは隣接列同士、または同一列の隣接行同士。
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
                # 高さが異なる隣接列は、列間で折れる 3 本の線をつなぎ、最後の線だけに矢印の先端を付ける。
                slide.line(start, cy1, mid, cy1, ACCENT, 1.5)
                slide.line(mid, cy1, mid, cy2, ACCENT, 1.5)
                slide.line(mid, cy2, end, cy2, ACCENT, 1.5, arrow=True)
        else:
            raise ValueError("矢印は隣の列、または同じ列の隣の行を結んでください")
    for node in nodes:
        # 接続線の後にカードを描くため、背景とラベルを線より手前に表示できる。
        _, _, x, y = places[node["id"]]
        slide.rect(x, y, w, row_height, PALE, RULE)
        put(slide, node["label"], x + 10, y + 4, w - 20, row_height - 8, 14,
            color=NAVY, bold=True, minimum=12, align="ctr", center=True)
    if page.get("note"):
        put(slide, page["note"], LEFT, BOTTOM - 23, CONTENT, 23, 11.5, color=MUTED)


# 工程表。左は作業名と担当、右は残りの幅を期間数で等分して作る。

# ========================================
# ロードマップのレイアウト
# ========================================
def roadmap(slide, page, _root):
    top = body_top(page)
    periods = page["periods"]
    tasks = page["tasks"]
    left_w = 208
    cell_w = (CONTENT - left_w) / len(periods)
    y = top + 12
    row_h = min(57, (306 - 30) / len(tasks))
    put(slide, "作業・担当", LEFT + 8, y, left_w - 16, 24, 13, color=DARK_GREEN, bold=True)
    for i, label in enumerate(periods):
        put(slide, label, LEFT + left_w + i * cell_w, y, cell_w, 24, 13,
            color=DARK_GREEN, bold=True, align="ctr")
    y += 33
    for i, task in enumerate(tasks):
        # 期間番号は 0 始まり。帯は start と end の両端を含む区画に描く。
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
        # 期間の境界線と重ならないよう、帯の左右を 4 ずつ短くする。
        bx = LEFT + left_w + start * cell_w + 4
        bw = (end - start + 1) * cell_w - 8
        slide.rect(bx, ry + 8, bw, row_h - 16, ACCENT)
        put(slide, task.get("outcome", ""), bx + 7, ry + 11, bw - 14, row_h - 22, 11,
            color=WHITE, minimum=10, center=True)
    rule(slide, y + len(tasks) * row_h)
    if page.get("note"):
        put(slide, page["note"], LEFT, BOTTOM - 38, CONTENT, 38, 12, color=MUTED)


# 左列に課題と影響を上下に重ね、中央に対応策、右に担当を示す。

# ========================================
# 課題と対応策のレイアウト
# ========================================
def issues(slide, page, _root):
    top = body_top(page)
    columns = [(LEFT + 11, 308), (LEFT + 351, 391), (RIGHT - 104, 93)]
    y = top + 9
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
    # 注記用の領域を除いた高さで、各行の最も高い列に合わせて行高を決める。文字サイズは 13 から試す。
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
    # 余った高さを各行に均等配分する。行ごとの増分は 8 以下とし、背景を交互に変えて行を追いやすくする。
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


# ========================================
# レイアウト名と描画関数の登録
# ========================================
# 入力の layout 名と描画関数を対応させる。形式を増やす場合は slide.schema.json にも追加する。
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
