"""対応する LaTeX 記法を、PowerPoint の編集可能な数式に変換する。"""

# ========================================
# 数式の要素と対応する命令
# ========================================
from dataclasses import dataclass
from math import isclose
import re
from xml.sax.saxutils import escape

from pptx_document import display_width


SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "eta": "η",
    "theta": "θ", "lambda": "λ", "mu": "μ", "pi": "π", "sigma": "σ",
    "phi": "φ", "omega": "ω", "Delta": "Δ", "Sigma": "Σ", "partial": "∂",
    "nabla": "∇", "infty": "∞", "cdot": "⋅", "times": "×", "pm": "±",
    "leq": "≤", "geq": "≥", "ne": "≠", "to": "→", "in": "∈",
    "log": "log", "ln": "ln", "exp": "exp", "max": "max", "min": "min",
    ",": " ", "{": "{", "}": "}", "%": "%",
}
MATRICES = {"matrix": ("", ""), "bmatrix": ("[", "]"), "pmatrix": ("(", ")")}
TOKEN = re.compile(r"\\[A-Za-z]+|\\[^A-Za-z]|[\s\S]")


@dataclass(frozen=True)
class Formula:
    xml: str
    plain: str
    width: float   # 本文の文字サイズを 1 とする概算幅
    height: float  # 基準行高を 1 とする概算高さ


def math_run(value, upright=False):
    style = '<m:rPr><m:sty m:val="p"/></m:rPr>' if upright else ""
    return Formula(
        f'<m:r>{style}<m:t xml:space="preserve">{escape(value)}</m:t></m:r>',
        value, display_width(value), 1,
    )


def joined(parts):
    return Formula(
        "".join(p.xml for p in parts), "".join(p.plain for p in parts),
        sum(p.width for p in parts), max(p.height for p in parts),
    )


def fraction(num, den):
    return Formula(
        f'<m:f><m:num>{num.xml}</m:num><m:den>{den.xml}</m:den></m:f>',
        f'({num.plain})/({den.plain})', max(num.width, den.width) + 0.7,
        0.8 * (num.height + den.height) + 0.3,
    )


def matrix(rows, left, right):
    columns = len(rows[0])
    if any(len(row) != columns for row in rows):
        raise ValueError("行列の列数をそろえてください")
    cells = "".join(
        '<m:mr>' + "".join(f'<m:e>{cell.xml}</m:e>' for cell in row) + '</m:mr>'
        for row in rows
    )
    width = sum(max(row[j].width for row in rows) for j in range(columns))
    height = sum(max(cell.height for cell in row) for row in rows)
    return Formula(
        f'<m:d><m:dPr><m:begChr m:val="{left}"/><m:endChr m:val="{right}"/></m:dPr>'
        f'<m:e><m:m>{cells}</m:m></m:e></m:d>',
        left + "; ".join(", ".join(cell.plain for cell in row) for row in rows) + right,
        width + 0.65 * columns + 0.6, 0.25 + 1.03 * height,
    )


# ========================================
# LaTeX 記法の読み取り
# ========================================
class Parser:
    def __init__(self, source):
        # LaTeX の空白は命令の区切りとして扱う。式中の空きは \, で指定できる。
        self.tokens = [token for token in TOKEN.findall(source) if not token.isspace()]
        self.index = 0

    def peek(self):
        return self.tokens[self.index] if self.index < len(self.tokens) else None

    def take(self):
        token = self.peek()
        if token is None:
            raise ValueError("数式が途中で終わっています")
        self.index += 1
        return token

    def expect(self, token):
        if self.take() != token:
            raise ValueError(f"数式に {token} が必要です")

    def group(self, depth):
        self.expect("{")
        return self.expression_group_body(depth)

    def environment(self):
        self.expect("{")
        name = []
        while self.peek() not in (None, "}"):
            name.append(self.take())
        self.expect("}")
        name = "".join(name)
        if name not in MATRICES:
            raise ValueError(f"未対応の数式環境です（{name}）")
        return name

    def read_matrix(self, depth):
        name = self.environment()
        rows = [[]]
        while True:
            rows[-1].append(self.expression({"&", r"\\", r"\end"}, depth))
            sep = self.take()
            if sep == r"\end":
                if self.environment() != name:
                    raise ValueError("行列の開始と終了の名前が一致しません")
                return matrix(rows, *MATRICES[name])
            if sep == r"\\":
                rows.append([])

    def expression(self, stops, depth=0):
        if depth > 24:
            raise ValueError("数式の入れ子が深すぎます")
        parts = []
        while self.peek() is not None and self.peek() not in stops:
            parts.append(self.term(depth))
        if not parts:
            raise ValueError("空の数式・添字・行列セルがあります")
        return joined(parts)

    def term(self, depth):
        token = self.take()
        base = self.atom(token, depth + 1)
        scripts = {}
        while self.peek() in ("_", "^"):
            sign = self.take()
            if sign in scripts:
                raise ValueError("同じ位置に添字が重複しています")
            scripts[sign] = self.atom(self.take(), depth + 1)
        if not scripts:
            return base
        below, above = scripts.get("_"), scripts.get("^")
        lower = f'<m:sub>{below.xml}</m:sub>' if below else ""
        upper = f'<m:sup>{above.xml}</m:sup>' if above else ""
        if token == r"\sum":
            lower = lower or '<m:sub/>'
            upper = upper or '<m:sup/>'
            properties = (
                '<m:subHide m:val="1"/>' if not below else ""
            ) + ('<m:supHide m:val="1"/>' if not above else "")
            # 直後の項を総和の内部に取り込むと、PowerPoint の数式としてまとまりを保てる。
            following = self.peek()
            operand = (
                self.term(depth + 1) if following is not None and depth < 24
                and (following[0].isalnum() or following in ("{", r"\frac", r"\begin"))
                else math_run("")
            )
            xml = ('<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="subSup"/>'
                   f'{properties}</m:naryPr>{lower}{upper}<m:e>{operand.xml}</m:e></m:nary>')
        else:
            operand = math_run("")
            tag = "sSubSup" if below and above else ("sSub" if below else "sSup")
            xml = f'<m:{tag}><m:e>{base.xml}</m:e>{lower}{upper}</m:{tag}>'
        script_width = max(part.width for part in scripts.values())
        extra_height = sum(0.4 * part.height for part in scripts.values())
        plain = base.plain + "".join(f'{sign}({part.plain})' for sign, part in scripts.items())
        return Formula(xml, plain + operand.plain, base.width + 0.85 * script_width + operand.width,
                       max(base.height + extra_height, operand.height))

    def atom(self, token, depth):
        if token == "{":
            return self.expression_group_body(depth)
        if token == r"\frac":
            return fraction(self.group(depth), self.group(depth))
        if token == r"\begin":
            return self.read_matrix(depth)
        if token == r"\sum":
            return math_run("∑", upright=True)
        if token.startswith("\\"):
            if token[1:] not in SYMBOLS:
                raise ValueError(f"未対応の数式命令です（{token}）")
            return math_run(SYMBOLS[token[1:]], upright=True)
        if token in ("}", "_", "^", "&"):
            raise ValueError(f"数式の記号の位置を確認してください（{token}）")
        return math_run("−" if token == "-" else token)

    def expression_group_body(self, depth):
        content = self.expression({"}"}, depth)
        self.expect("}")
        return content


def parse_equation(source):
    """数式を検証し、ネイティブ数式と代替テキストを描画できる形にする。"""
    return Parser(source).expression(set())


def equation_size(formula, width):
    # 行列や分数は折り返せない。下限を下回る長い式は位置付きのエラーにする。
    desired = 19
    size = min(desired, width / (1.22 * formula.width)) if formula.width else desired
    if size < 12 and not isclose(size, 12):
        raise ValueError("数式が幅に収まりません。全幅に移すか、式を分けてください")
    return round(size, 1)
