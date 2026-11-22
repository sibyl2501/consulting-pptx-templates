# スライドの図形・表・テーマを Office Open XML（OOXML）で記述し、ZIP 形式の PPTX にまとめる。
# レイアウトが渡す座標・寸法・文字サイズは基準キャンバスでの値。書き出し時に用紙に合わせて換算する。
"""Write editable PowerPoint shapes, tables and an Office theme using only Python."""

# ========================================
# ライブラリの読み込み
# ========================================
from itertools import count
from math import ceil, isfinite
from os import replace
from tempfile import NamedTemporaryFile
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZIP_DEFLATED, ZipFile
import unicodedata

# ========================================
# キャンバスと用紙サイズの指定
# ========================================
# layouts.py が使う基準キャンバス。座標の原点は左上で、x は右、y は下に増える。
WIDTH, HEIGHT = 960, 540
# 出力サイズはポイント（1 インチ = 72 pt）。A4 横は 297 × 210 mm。
DEFAULT_PAGE_SIZE = "a4"
PAGE_SIZES = {
    "a4": (297 * 72 / 25.4, 210 * 72 / 25.4, "A4"),
    "16:9": (WIDTH, HEIGHT, "screen16x9"),
}

DEFAULT_HEADING_SIZE = 22
MIN_HEADING_SIZE = 20
MAX_HEADING_SIZE = 24


def canvas_scale(page_size):
    return PAGE_SIZES[page_size][0] / WIDTH


def canvas_height(page_size):
    """横幅を基準に縦横等倍で描くため、用紙の高さをレイアウト座標に戻す。"""
    return PAGE_SIZES[page_size][1] / canvas_scale(page_size)


# ========================================
# フォントの指定
# ========================================
# テキストの直接指定と、テーマの既定フォントで共通に使うフォント名。
DEFAULT_FONT = "hiragino"
FONT_FAMILIES = {
    "hiragino": "Hiragino Sans",
    "yu-gothic": "Yu Gothic",
}

# ========================================
# XMLの名前空間と共通設定
# ========================================
# DrawingML の図形・文字・表要素に使う a: 名前空間。
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
# graphicData に渡す表の種類識別子。ここで PowerPoint が中身を表と判定する。
TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"
# PresentationML のスライド・図形コンテナなどに使う p: 名前空間。
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
# r:id と、各 .rels 内の関係タイプの URI に使う Office 文書用の名前空間。
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# プレゼンテーション、スライド、マスターなどのルート要素に共通で付ける XML 名前空間宣言。
NS = f'xmlns:a="{A}" xmlns:p="{P}" xmlns:r="{R}"'
# ZIP 内に保存する各 XML 部品の先頭に付ける宣言。
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
# 図形ツリーの根に必要なグループ属性。ID 1 を使うため、個別の図形・表は 2 から採番する。
GROUP = ('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
         '<p:grpSpPr/>')
# マスターの配色対応表。背景 bg と文字 tx の番号を、テーマ側の明色 lt・暗色 dk に対応させる。
COLOR_MAP = ('accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
             'accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" tx1="dk1" '
             'tx2="dk2" hlink="hlink" folHlink="folHlink"')
# スライドや白紙レイアウトで、マスター側の色の対応付けを引き継ぐ。
INHERIT = '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'


# ========================================
# 座標・色・線の変換
# ========================================
# ポイントを OOXML の長さ単位 EMU（English Metric Unit）に変換する。1 pt は 12,700 EMU。
def emu(points):
    return round(points * 12700)


# 位置 (x, y) とサイズ (w, h) を変換要素にする。tag で図形用と表用を選び、attrs で反転も渡せる。
def transform(tag, x, y, w, h, attrs=""):
    return (f'<{tag}{attrs}><a:off x="{emu(x)}" y="{emu(y)}"/>'
            f'<a:ext cx="{emu(w)}" cy="{emu(h)}"/></{tag}>')


# RGB の16進文字列（# なし）で単色塗りを指定する。None は透明な塗りとして扱う。
def fill(color):
    if color is None:
        return '<a:noFill/>'
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


# 線・枠の色と太さを OOXML に変換する。色を省略した場合は線が透明になる。
def outline(color=None, width=0.75, arrow=False):
    tail = '<a:tailEnd type="triangle" w="sm" len="sm"/>' if arrow else ""
    return f'<a:ln w="{emu(width)}">{fill(color)}{tail}</a:ln>'


# ========================================
# 文字の書式・幅・高さとXMLの生成
# ========================================
# 欧文、東アジア文字、複合文字用の各スロットに同じ書体名を指定する。
def fonts(font, latin_font=None):
    face = FONT_FAMILIES[font]
    latin = latin_font or face
    return f'<a:latin typeface="{latin}"/><a:ea typeface="{face}"/><a:cs typeface="{latin}"/>'


# 文字列の一部分（ラン）のサイズ、色、太字、フォントを指定する。OOXML のサイズは 1/100 pt 単位。
def font_properties(size, color, bold, font, latin_font=None):
    return (f'<a:rPr lang="ja-JP" sz="{round(size * 100)}" b="{int(bold)}">'
            f'{fill(color)}{fonts(font, latin_font)}</a:rPr>')


def display_width(value):
    """日本語と欧文を含む行の幅を em 単位で概算する。"""
    return sum(
        0 if unicodedata.combining(char) else
        1.0 if unicodedata.east_asian_width(char) in "WF" else
        0.32 if char.isspace() else
        0.46 if char in "ilI.,:;'|" else
        0.76 if char in "MW@%mw" else 0.61
        for char in value
    )


def text_height(value, width, size, bold=False):
    """禁則処理や代替フォントの差を見込んで、高さを保守的に見積もる。"""
    capacity = width / size / (1.12 if bold else 1.07)
    lines = sum(max(1, ceil(display_width(line) / capacity)) for line in str(value).split("\n"))
    return lines * size * 1.30


def cell_text(value):
    return "—" if value is None else str(value)


# 1 行を 1 段落にする。align は l（左）・ctr（中央）・r（右）で、行間は指定値の 116%。
# escape() で本文中の & や < を XML 用に処理し、xml:space で空白を保持する。
def paragraph(value, size, color, bold, align, font, latin_font=None):
    return (f'<a:p><a:pPr algn="{align}"><a:lnSpc><a:spcPct val="116000"/>'
            f'</a:lnSpc></a:pPr><a:r>{font_properties(size, color, bold, font, latin_font)}'
            f'<a:t xml:space="preserve">{escape(value)}</a:t></a:r>'
            f'<a:endParaRPr lang="ja-JP" sz="{round(size * 100)}"/></a:p>')


def fitted_text_size(value, w, h, size, bold, minimum=None):
    """文字サイズを指定下限まで調整し、領域に収まるサイズを返す。"""
    minimum = size if minimum is None else minimum
    if w <= 0 or h <= 0 or minimum > size:
        raise ValueError("文字領域または最小文字サイズの指定が不正です")
    while value and text_height(value, w, size, bold) > h + 0.1:
        if size <= minimum:
            snippet = value.replace("\n", " ")[:44]
            raise ValueError(f"文字が収まりません（下限 {minimum:g} pt）『{snippet}』")
        size = max(minimum, size - 0.5)
    return size


def text_body(tag, value, w, h, size, color, bold, align, anchor, minimum=None, *, scale=1, font=DEFAULT_FONT):
    value = str(value)
    anchor = {"middle": "ctr", "top": "t"}.get(anchor, anchor)
    size = fitted_text_size(value, w, h, size, bold, minimum)
    body = (f'<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}">'
            '<a:noAutofit/></a:bodyPr>')
    return (f'<{tag}>{body}<a:lstStyle/>' +
            ''.join(paragraph(line, size * scale, color, bold, align, font) for line in value.split("\n")) +
            f'</{tag}>')


# ========================================
# スライドの組み立て
# ========================================
# 1 枚のスライドを組み立てる。図形や表の XML を追加順に保持し、後から追加した要素が手前に描画される。
class Slide:
    # スライド名、要素の格納先、スライド内で一意な図形 ID の採番を初期化する。
    def __init__(self, name, page_size=DEFAULT_PAGE_SIZE, *, font=DEFAULT_FONT, heading_size=DEFAULT_HEADING_SIZE):
        self.name = name
        self.page_size = page_size
        self.font = font
        self.heading_size = heading_size
        self.height = canvas_height(page_size)
        self.x_scale = self.y_scale = canvas_scale(page_size)
        self.parts = []
        self.ids = count(2)
        self.pictures = []
        self.warnings = []
        self.readable_regions = []

    # 背景や罫線と本文の重なりは意図した装飾なので、文字・数式・表などを対象にする。
    def track_region(self, label, x, y, w, h):
        for other, left, top, width, height in self.readable_regions:
            horizontal = min(x + w, left + width) - max(x, left)
            vertical = min(y + h, top + height) - max(y, top)
            if horizontal > 3 and vertical > 3:
                self.warnings.append(f"「{label}」と「{other}」が重なる可能性があります")
        self.readable_regions.append((label, x, y, w, h))

    def inspect_text(self, label, value, x, y, w, h, size, bold, align, anchor):
        if not value.strip():
            return
        # 文字幅の概算を使う。日本語は文字境界、欧文は単語の途中を避けて折り返す。
        capacity = w / size / (1.12 if bold else 1.07)
        widths = []
        for paragraph_text in value.split("\n"):
            line, remainder = "", paragraph_text
            wrapped = False
            while remainder and display_width(remainder) > capacity:
                index = 0
                width = 0
                for index, char in enumerate(remainder):
                    width += display_width(char)
                    if width > capacity:
                        break
                cutoff = max(1, index)
                boundary = remainder.rfind(" ", 0, cutoff + 1)
                if boundary > 0 and remainder[boundary + 1:cutoff].isascii():
                    cutoff = boundary
                line, remainder = remainder[:cutoff], remainder[cutoff:].lstrip()
                widths.append(display_width(line.rstrip()))
                wrapped = True
            widths.append(display_width(remainder.rstrip()))
            if wrapped and 0 < len(remainder.strip()) <= 2:
                self.warnings.append(f"「{label}」は最終行が短くなる可能性があります")
        used_w = min(w, max(widths) * size)
        used_h = min(h, len(widths) * size * 1.16 * self.x_scale / self.y_scale)
        if align == "ctr":
            x += (w - used_w) / 2
        elif align == "r":
            x += w - used_w
        if anchor in ("middle", "ctr"):
            y += (h - used_h) / 2
        self.track_region(label, x, y, used_w, used_h)

    def check_bounds(self, name, x, y, w, h):
        if (not all(isfinite(value) for value in (x, y, w, h))
                or min(x, y, w, h) < 0 or x + w > WIDTH + 0.01 or y + h > self.height + 0.01):
            raise ValueError(f"スライドの外に配置されています（{name}）")

    def fit_picture(self, pixel_width, pixel_height, frame_width, frame_height):
        """用紙上で元の縦横比を保つ画像の幅・高さを、基準キャンバスの単位で返す。"""
        scale = min(frame_width * self.x_scale / pixel_width,
                    frame_height * self.y_scale / pixel_height)
        return pixel_width * scale / self.x_scale, pixel_height * scale / self.y_scale

    def _transform(self, tag, x, y, w, h, attrs="", *, keep_ratio=False):
        """基準キャンバスの座標を、選んだ用紙の寸法に変換する。"""
        x, y, w, h = x * self.x_scale, y * self.y_scale, w * self.x_scale, h * self.y_scale
        if keep_ratio:
            # 円などの縦横比を保ち、変換前の図形の中心が同じ位置に来るよう補正する。
            scale = min(self.x_scale, self.y_scale)
            scaled_w, scaled_h = w * scale / self.x_scale, h * scale / self.y_scale
            x, y = x + (w - scaled_w) / 2, y + (h - scaled_h) / 2
            w, h = scaled_w, scaled_h
        return transform(tag, x, y, w, h, attrs)

    # 矩形・線・テキストに共通する図形 XML を追加し、基準キャンバス内に収まるか確認する。
    def _shape(self, name, x, y, w, h, geometry, color, border, body="", flips=""):
        self.check_bounds(name, x, y, w, h)
        # テキストを渡された図形には txBox を指定する。quoteattr() は図形名を XML 属性用に引用・エスケープする。
        shape_id = next(self.ids)
        textbox = ' txBox="1"' if body else ''
        self.parts.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name={quoteattr(name)}/>'
            f'<p:cNvSpPr{textbox}/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'{self._transform("a:xfrm", x, y, w, h, flips, keep_ratio=geometry == "ellipse")}'
            f'<a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom>'
            f'{fill(color)}{border}</p:spPr>{body}</p:sp>'
        )

    # 背景・罫線などの図形を配置する。geometry で rect、ellipse、rightArrow などの形状を指定できる。
    def rect(self, x, y, w, h, color, border=None, name="背景・罫線", geometry="rect"):
        self._shape(name, x, y, w, h, geometry, color, outline(border, 0.75 * self.x_scale))

    # 始点・終点から線を描く。小さい側の座標と正の寸法を取り、端点の向きは水平・垂直反転で表す。
    def line(self, x1, y1, x2, y2, color, width=1, name="線", arrow=False):
        flip_h = ' flipH="1"' if x2 < x1 else ''
        flip_v = ' flipV="1"' if y2 < y1 else ''
        self._shape(name, min(x1, x2), min(y1, y2), abs(x1-x2), abs(y1-y2),
                    "line", None, outline(color, width * self.x_scale, arrow), flips=flip_h + flip_v)

    # 指定位置に編集可能なテキストボックスを作る。name を省くと本文の先頭を図形名に使う。
    def text(self, value, x, y, w, h, size, color, bold=False, align="l",
             anchor="top", name=None, minimum=None):
        value = str(value)
        label = name or value.replace("\n", " ")[:50] or "本文"
        try:
            # 横幅と文字を同じ倍率で換算し、基準キャンバスで見積もった改行を保ちやすくする。
            fitted = fitted_text_size(value, w, h, size, bold, minimum)
            content = text_body("p:txBody", value, w, h, fitted, color, bold, align, anchor,
                                scale=self.x_scale, font=self.font)
        except ValueError as error:
            raise ValueError(f"{label} — {error}") from error
        self._shape(label, x, y, w, h, "rect", None, outline(), content)
        if fitted < size:
            self.warnings.append(f"「{label}」の文字サイズを "
                                 f"{size * self.x_scale:.1f} → {fitted * self.x_scale:.1f} pt に縮小しました")
        self.inspect_text(label, value, x, y, w, h, fitted, bold, align, anchor)

    def code(self, lines, x, y, w, h, size, color, label):
        """コードは等幅で描き、自動折り返しを無効にする。"""
        body = ('<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">'
                '<a:noAutofit/></a:bodyPr><a:lstStyle/>' +
                ''.join(paragraph(line, size * self.x_scale, color, False, "l", self.font, "Consolas")
                        for line in lines) + '</p:txBody>')
        self._shape(label, x, y, w, h, "rect", None, outline(), body)
        self.track_region(label, x, y, w, h)

    def equation(self, formula, x, y, w, h, size, color, label, *, desired=19, align="ctr"):
        """ネイティブ数式を使えない閲覧ソフト向けに、同じ位置に代替テキストを入れる。"""
        fallback = text_body("p:txBody", formula.plain, w, h, min(12, size), color,
                             False, align, "middle", 8, scale=self.x_scale, font=self.font)
        math_font = round(size * self.x_scale * 100)
        style = (f'<a:rPr lang="en-US" sz="{math_font}">{fill(color)}'
                 '<a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:rPr>')
        # m:rPr（数学用の書式）と m:t（文字）の間に DrawingML の書体を配置する。
        native = formula.xml.replace("<m:t ", f"{style}<m:t ")
        justification = "left" if align == "l" else "center"
        body = ('<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr">'
                f'<a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="{align}"/>'
                f'<a14:m xmlns:m="{M}"><m:oMathPara><m:oMathParaPr><m:jc m:val="{justification}"/>'
                f'</m:oMathParaPr><m:oMath>{native}</m:oMath></m:oMathPara></a14:m>'
                f'<a:endParaRPr lang="en-US" sz="{math_font}"/></a:p></p:txBody>')
        self._shape(label, x, y, w, h, "rect", None, outline(), fallback)
        fallback_shape = self.parts.pop()
        native_shape = fallback_shape.replace(fallback, body, 1)
        self.parts.append(f'<mc:AlternateContent xmlns:mc="{MC}">'
                          f'<mc:Choice xmlns:a14="{A14}" Requires="a14">{native_shape}</mc:Choice>'
                          f'<mc:Fallback>{fallback_shape}</mc:Fallback></mc:AlternateContent>')
        self.track_region(label, x, y, w, h)
        if size < desired - 0.05:
            self.warnings.append(f"「{label}」の数式を {math_font / 100:.1f} pt に縮小しました")

    def picture(self, data, extension, x, y, w, h, alternative):
        """元画像を貼り、画像ボックスと注釈を個別に編集できるようにする。"""
        self.check_bounds("画像", x, y, w, h)
        rid = len(self.pictures) + 2  # rId1 は白紙レイアウト。
        shape_id = next(self.ids)
        self.pictures.append((data, extension))
        self.parts.append(
            f'<p:pic><p:nvPicPr><p:cNvPr id="{shape_id}" name="図版" descr={quoteattr(alternative)}/>'
            '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="rId{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{self._transform("a:xfrm", x, y, w, h)}'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        )

    # 列幅・行高・セル値から編集可能な PowerPoint の表を作る。highlight は強調する列の 0 始まりの番号。
    def table(self, x, y, widths, heights, rows, navy, accent_light, ink, highlight=None, size=12, aligns=()):
        """A real PowerPoint table. Each cell remains directly editable."""
        self.check_bounds("表", x, y, sum(widths), sum(heights))
        self.track_region("表", x, y, sum(widths), sum(heights))
        tbl_rows = []
        # strict=True によって、行数や列数が指定した寸法と合わない入力を検出する。
        for row_index, (height, values) in enumerate(zip(heights, rows, strict=True)):
            cells = []
            for col_index, (width, value) in enumerate(zip(widths, values, strict=True)):
                # 強調列と交互の行背景を設定する。先頭行は見出しなので濃紺背景と白文字を優先する。
                bg = accent_light if col_index == highlight else (
                    "F2F5F8" if col_index == 0 or row_index % 2 == 0 else "FFFFFF"
                )
                color = ink
                if row_index == 0:
                    bg, color = navy, "FFFFFF"
                elif col_index == 0:
                    color = navy
                # セルの 4 辺に同じ 0.65 pt の罫線を設定する。
                borders = ''.join(
                    f'<a:ln{edge} w="{emu(0.65 * self.x_scale)}">{fill("D8E2EB")}</a:ln{edge}>'
                    for edge in ("L", "R", "T", "B")
                )
                align = aligns[col_index] if aligns and row_index else "l"
                content = text_body("a:txBody", cell_text(value), width - 16, height - 8, size,
                                    color, row_index == 0 or col_index == 0, align, "middle",
                                    scale=self.x_scale, font=self.font)
                cells.append(f'<a:tc>{content}<a:tcPr marL="{emu(8 * self.x_scale)}" '
                             f'marR="{emu(8 * self.x_scale)}" marT="{emu(4 * self.y_scale)}" '
                             f'marB="{emu(4 * self.y_scale)}" anchor="ctr">'
                             f'{borders}{fill(bg)}</a:tcPr></a:tc>')
            tbl_rows.append(f'<a:tr h="{emu(height * self.y_scale)}">{"".join(cells)}</a:tr>')
        # 列のグリッド定義と全体の外枠を作り、各セルを持つ表として図形ツリーに追加する。
        grid = ''.join(f'<a:gridCol w="{emu(width * self.x_scale)}"/>' for width in widths)
        shape_id = next(self.ids)
        self.parts.append(
            f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="編集用の表"/>'
            '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            f'{self._transform("p:xfrm", x, y, sum(widths), sum(heights))}'
            f'<a:graphic><a:graphicData uri="{TABLE}"><a:tbl><a:tblPr/><a:tblGrid>'
            f'{grid}</a:tblGrid>{"".join(tbl_rows)}</a:tbl></a:graphicData></a:graphic>'
            '</p:graphicFrame>'
        )

    # 保持している図形と表を描画順に結合し、スライド 1 枚分の XML を返す。
    def xml(self):
        return (f'<p:sld {NS}><p:cSld name={quoteattr(self.name)}><p:spTree>{GROUP}'
                f'{"".join(self.parts)}</p:spTree></p:cSld>{INHERIT}</p:sld>')


# ========================================
# XML部品間の参照関係
# ========================================
# (関係の種類, 相対パス) の組から .rels を作る。引数の順に rId1 から割り当てる。
def relationships(*items):
    links = ''.join(
        f'<Relationship Id="rId{i}" Type="{R}/{kind}" Target={quoteattr(target)}/>'
        for i, (kind, target) in enumerate(items, 1)
    )
    return f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{links}</Relationships>'


# ========================================
# テーマの色とフォント
# ========================================
# Office のテーマ XML を作る。テーマ色は PowerPoint 上で追加する図形などの色選択にも使われる。
def theme(navy, accent, font=DEFAULT_FONT):
    # dk / lt は基本の暗色・明色、accent は強調色、hlink / folHlink は未訪問・訪問済みリンクの色。
    palette = {
        "dk1": "182D41", "lt1": "FFFFFF", "dk2": navy, "lt2": "F1F5F8",
        "accent1": accent, "accent2": "6C92B3", "accent3": "A45A29",
        "accent4": "789080", "accent5": "69739B", "accent6": "D1A156",
        "hlink": "14629B", "folHlink": "694E83",
    }
    colors = ''.join(f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>' for k, v in palette.items())
    # 見出し用（major）と本文用（minor）の両方に選んだ書体を入れ、日本語（Jpan）も明示する。
    font_scheme = ''.join(
        f'<a:{slot}Font>{fonts(font)}<a:font script="Jpan" typeface="{FONT_FAMILIES[font]}"/></a:{slot}Font>'
        for slot in ("major", "minor")
    )
    # テーマの塗り・線に使用する色の参照と、線の既定の太さ 3 段階（EMU）を用意する。
    default_fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    lines = ''.join(f'<a:ln w="{w}">{default_fill}<a:prstDash val="solid"/></a:ln>'
                    for w in (9525, 19050, 28575))
    return (f'<a:theme xmlns:a="{A}" name="Consulting Slides"><a:themeElements>'
            f'<a:clrScheme name="Navy">{colors}</a:clrScheme><a:fontScheme name="Japanese">'
            f'{font_scheme}</a:fontScheme><a:fmtScheme name="Standard">'
            f'<a:fillStyleLst>{default_fill * 3}</a:fillStyleLst><a:lnStyleLst>{lines}</a:lnStyleLst>'
            f'<a:effectStyleLst>{"<a:effectStyle><a:effectLst/></a:effectStyle>" * 3}</a:effectStyleLst>'
            f'<a:bgFillStyleLst>{default_fill * 3}</a:bgFillStyleLst>'
            '</a:fmtScheme></a:themeElements></a:theme>')


# ========================================
# PowerPointファイルの書き出し
# ========================================
# Slide の一覧とテーマ色を受け取り、編集可能な .pptx を指定パスへ保存する。
def write_pptx(path, slides, navy, accent):
    # PPTX 内のページ寸法は資料全体に共通。異なるサイズのスライドを混在させない。
    page_size = slides[0].page_size if slides else DEFAULT_PAGE_SIZE
    if any(slide.page_size != page_size for slide in slides):
        raise ValueError("資料内のスライドサイズをそろえてください")
    paper_width, paper_height, size_type = PAGE_SIZES[page_size]
    font = slides[0].font if slides else DEFAULT_FONT
    if any(slide.font != font for slide in slides):
        raise ValueError("資料内の書体の設定をそろえてください")
    presentation = "presentationml."
    parts = {
        # パッケージを開いたときの入口。ここからプレゼンテーション本体にたどる。
        "_rels/.rels": relationships(("officeDocument", "ppt/presentation.xml")),
        # 全体設定とスライドの表示順。マスターに rId1、本文・表紙を含む各スライドに rId2 以降を割り当てる。
        "ppt/presentation.xml": (
            f'<p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/>'
            f'</p:sldMasterIdLst><p:sldIdLst>'
            + ''.join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(len(slides)))
            + f'</p:sldIdLst><p:sldSz cx="{emu(paper_width)}" cy="{emu(paper_height)}" type="{size_type}"/>'
            # ノートのページサイズは EMU で指定し、文書の既定テキスト言語を日本語にする。
            '<p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr>'
            '<a:defRPr lang="ja-JP"/></a:defPPr></p:defaultTextStyle></p:presentation>'
        ),
        # 本体から各部品へのリンク。relationships() の引数順は、本体側の r:id に合わせる。
        "ppt/_rels/presentation.xml.rels": relationships(
            ("slideMaster", "slideMasters/slideMaster1.xml"),
            *(("slide", f"slides/slide{i}.xml") for i in range(1, len(slides)+1)),
        ),
        "ppt/theme/theme1.xml": theme(navy, accent, font),
        # 全スライドが使う白地のマスター。配色の対応付けと白紙レイアウトへの参照を持つ。
        "ppt/slideMasters/slideMaster1.xml": (
            f'<p:sldMaster {NS}><p:cSld name="共通マスター"><p:bg><p:bgPr>{fill("FFFFFF")}'
            f'</p:bgPr></p:bg><p:spTree>{GROUP}</p:spTree></p:cSld><p:clrMap {COLOR_MAP}/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
            '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'
        ),
        # 共通マスターから白紙レイアウトとテーマへのリンクを持たせる。
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": relationships(
            ("slideLayout", "../slideLayouts/slideLayout1.xml"), ("theme", "../theme/theme1.xml")
        ),
        # マスターを参照する白紙レイアウト。個々の配置は各スライドの図形ツリーで決める。
        "ppt/slideLayouts/slideLayout1.xml": (
            f'<p:sldLayout {NS} type="blank" preserve="1"><p:cSld name="白紙">'
            f'<p:spTree>{GROUP}</p:spTree></p:cSld>{INHERIT}</p:sldLayout>'
        ),
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": relationships(
            ("slideMaster", "../slideMasters/slideMaster1.xml")
        ),
    }
    # XML 部品ごとの種類を登録する。後で [Content_Types].xml の上書き定義に展開する。
    types = {
        "ppt/presentation.xml": presentation + "presentation.main",
        "ppt/theme/theme1.xml": "theme",
        "ppt/slideMasters/slideMaster1.xml": presentation + "slideMaster",
        "ppt/slideLayouts/slideLayout1.xml": presentation + "slideLayout",
    }
    # ページごとの XML と、各ページから共通の白紙レイアウトへのリンクを追加する。
    for index, slide in enumerate(slides, 1):
        target = f"ppt/slides/slide{index}.xml"
        parts[target] = slide.xml()
        types[target] = presentation + "slide"
        links = [("slideLayout", "../slideLayouts/slideLayout1.xml")]
        for picture_index, (data, extension) in enumerate(slide.pictures, 1):
            filename = f"image{index}-{picture_index}.{extension}"
            parts[f"ppt/media/{filename}"] = data
            links.append(("image", f"../media/{filename}"))
        parts[f"ppt/slides/_rels/slide{index}.xml.rels"] = relationships(*links)
    # 各専用 XML の ContentType を列挙する。拡張子既定値だけでは区別できない部品の種類を宣言する。
    overrides = ''.join(f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-'
                        f'officedocument.{kind}+xml"/>' for part, kind in types.items())
    parts["[Content_Types].xml"] = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="jpg" ContentType="image/jpeg"/>'
        f'<Default Extension="xml" ContentType="application/xml"/>{overrides}</Types>'
    )
    # 出力先の親ディレクトリを用意し、すべての部品を圧縮して 1 個の PPTX にまとめる。
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, prefix=f".{path.name}.", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as package:
            for name, contents in parts.items():
                package.writestr(name, XML + contents if isinstance(contents, str) else contents)
        replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)

