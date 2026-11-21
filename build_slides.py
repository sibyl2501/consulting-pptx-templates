#!/usr/bin/env python3
"""閲覧用のスライドを JSON から生成する。Python 3.10 以降・標準ライブラリのみ。"""

# ========================================
# ライブラリの読み込み
# ========================================
import argparse
import json
import math
from pathlib import Path
import re

# このファイルは入力を検査し、layouts.py にページ別の描画を依頼する。
# pptx_document.py は図形や表をまとめて PPTX に書き出す。
from layouts import ACCENT, NAVY, RENDERERS, footer, standard_slide, table_pages
from pptx_document import DEFAULT_FONT, DEFAULT_PAGE_SIZE, FONT_FAMILIES, PAGE_SIZES, Slide, write_pptx


# ========================================
# 入力仕様の場所と使用できない文字の指定
# ========================================
# 起動した場所に関係なく、このスクリプトの隣に置かれた入力仕様を使う。
SCHEMA_PATH = Path(__file__).with_name("slide.schema.json")
# PPTX 内の XML に書けない制御文字や単独のサロゲートを検出する。
# XML で使えるタブ・改行・復帰は対象から外す。
XML_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


# ========================================
# 入力エラーの定義
# ========================================
class InputError(ValueError):
    """資料の入力箇所が分かるエラー。"""


# ========================================
# 入力仕様の読み込みとレイアウト一覧
# ========================================
# JSON の入力仕様を、検査に使う Python の辞書として読み込む。
def read_schema():
    return json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))


def catalog(schema):
    # oneOf はレイアウトの候補一覧。例えば "#/$defs/layout_cover" は共通定義の layout_cover を指す。
    # 定義内で layout に指定された固定値をキーにし、型名から仕様を引けるようにする。
    layout_refs = schema["properties"]["slides"]["items"]["oneOf"]
    definitions = schema["$defs"]
    formats = [definitions[item["$ref"].split("/")[-1]] for item in layout_refs]
    return {entry["properties"]["layout"]["const"]: entry for entry in formats}


# ========================================
# 入力データの型と値の検査
# ========================================
# Python では isinstance(True, int) が真になる。JSON の整数と真偽値を区別するため型を直接比較する。
# number では大きすぎる整数や、無限大などの有限でない小数も除く。
def matches_type(value, kind):
    if kind == "number":
        return (type(value) is int and abs(value) < 10 ** 300
                or type(value) is float and math.isfinite(value))
    return type(value) is {
        "integer": int, "string": str, "array": list, "object": dict,
        "boolean": bool, "null": type(None),
    }[kind]


def validate(value, rule, at, definitions):
    """このキットのスキーマが使う型・必須項目・値の範囲を検査する。"""
    # at はエラーに表示する入力位置。辞書や配列の中身を調べる際に、キー名や番号を追記する。
    # $ref があれば、このキットの共通定義を参照先から取り出して検査する。
    if "$ref" in rule:
        rule = definitions[rule["$ref"].split("/")[-1]]
    # type が一覧なら、いずれか一つに当てはまればよい。表のセルでは文字列や数値のほか null なども使える。
    kind = rule.get("type")
    if kind:
        expected = kind if isinstance(kind, list) else [kind]
        if not any(matches_type(value, name) for name in expected):
            raise InputError(f"{at} — {' / '.join(expected)} を指定してください")
    # const は決まった値、enum は選択肢の一覧。例えばグラフの種類は指定された選択肢から選ぶ。
    if "const" in rule and value != rule["const"]:
        raise InputError(f"{at} — {rule['const']} を指定してください")
    if "enum" in rule and value not in rule["enum"]:
        raise InputError(f"{at} — 次から選択してください: {', '.join(rule['enum'])}")
    if isinstance(value, str) and ("minLength" in rule and not value.strip()
                                  or XML_CONTROL.search(value)):
        raise InputError(f"{at} — 空白・使用できない制御文字の指定を確認してください")
    # 各比較は「許容範囲から外れたら真」。exclusiveMinimum は下限と同じ値も除く。
    for name, predicate, label in (
        ("minimum", lambda limit: value < limit, "以上"),
        ("maximum", lambda limit: value > limit, "以下"),
        ("exclusiveMinimum", lambda limit: value <= limit, "より大きく"),
    ):
        if name in rule and predicate(rule[name]):
            raise InputError(f"{at} — {rule[name]} {label}にしてください")
    # 配列の件数を調べた後、各要素を再帰的に検査する。エラー表示上の要素番号は 1 始まり。
    if kind == "array":
        count = len(value)
        if count < rule.get("minItems", 0):
            raise InputError(f"{at} — {rule['minItems']} 件以上を指定してください")
        if count > rule.get("maxItems", count):
            raise InputError(f"{at} — {rule['maxItems']} 件以内にしてください")
        for i, item in enumerate(value, 1):
            validate(item, rule["items"], f"{at}[{i}]", definitions)
    # 必須キーを確認してから、入力済みのキーと値を調べる。仕様にないキーも拒否し、名前の打ち間違いを拾う。
    if kind == "object":
        props = rule["properties"]
        missing = [name for name in rule["required"] if name not in value]
        if missing:
            raise InputError(f"{at} — 必須項目が不足しています: {', '.join(missing)}")
        for key, item in value.items():
            if key not in props:
                raise InputError(f"{at}.{key} — この項目は定義されていません")
            validate(item, props[key], f"{at}.{key}", definitions)


# ========================================
# 資料全体と各ページの検査
# ========================================
def validate_document(document, schema):
    kinds = catalog(schema)
    # レイアウトを追加するときは、JSON の入力仕様と Python の描画関数の両方に同じ型を登録する。
    if kinds.keys() != RENDERERS.keys():
        raise RuntimeError("スキーマと描画関数のレイアウト一覧をそろえてください")
    # レイアウト名で分岐を確定し、oneOf の長いエラー一覧を1ページの指摘に絞る。
    header = dict(schema, properties={
        "project": schema["properties"]["project"],
        "slides": {"type": "array", "minItems": 1, "items": {}},
    })
    # 先に資料全体の構造と slides の件数を検査し、その後は各ページを layout に対応した仕様で調べる。
    validate(document, header, "資料", schema["$defs"])
    for i, page in enumerate(document["slides"], 1):
        at = f"入力 {i} 枚目"
        if not isinstance(page, dict):
            raise InputError(f"{at} — object を指定してください")
        name = page.get("layout")
        if not isinstance(name, str) or name not in kinds:
            raise InputError(f"{at}.layout — --catalog にあるレイアウト名を指定してください")
        validate(page, kinds[name], f"{at}（{name}）", schema["$defs"])


# ========================================
# スライドの生成
# ========================================
def render(document, asset_root, schema=None, *, page_size=DEFAULT_PAGE_SIZE, font=DEFAULT_FONT):
    """書き出しと入力検査で同じ配置を実行する。"""
    # 通常は同梱の入力仕様を読み込む。呼び出し元から仕様を渡すこともできる。
    schema = read_schema() if schema is None else schema
    validate_document(document, schema)
    slides = []
    for i, page in enumerate(document["slides"], 1):
        kind = page["layout"]
        try:
            # 詳細表の入力 1 件が複数枚に展開することがある。
            # i は元の入力番号、number は分割後の PowerPoint の通し番号。
            records = table_pages(page) if kind == "data_table" else [page]
            for part in records:
                number = len(slides) + 1
                # 表紙と章扉は独自の構成。章扉にフッターを付け、本文には共通見出しとフッターを付ける。
                if kind in ("cover", "section"):
                    slide = Slide(f"{number:02d} {part['title']}", page_size, font=font)
                    if kind == "section":
                        footer(slide, part, number)
                else:
                    slide = standard_slide(part, document, number, page_size, font=font)
                # 型名に対応する関数を呼び、作成済みの Slide に本文を加える。
                # asset_root は画像を使う型で、相対パスの起点になる。
                RENDERERS[kind](slide, part, asset_root)
                slides.append(slide)
        # 配置エラーに元の入力番号、型、タイトルを添え、修正すべき入力を明示する。
        except ValueError as error:
            raise InputError(f"入力 {i} 枚目（{kind}、{page['title']}） — {error}") from error
    return slides


# ========================================
# コマンドの受付と実行
# ========================================
def main(argv=None):
    # argv を省略するとターミナルの引数を読む。呼び出し元が引数の一覧を明示してもよい。
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, nargs="?", help="入力するJSON")
    parser.add_argument("output", type=Path, nargs="?", help="生成するPPTX")
    # 用紙指定は資料全体に適用する。既定は印刷しやすい A4 横。
    parser.add_argument("--page-size", choices=PAGE_SIZES, default=DEFAULT_PAGE_SIZE,
                        help="スライドの大きさ（既定: A4 横）")
    parser.add_argument("--font", choices=FONT_FAMILIES, default=DEFAULT_FONT,
                        help="書体（既定: ヒラギノ角ゴシック、代替: yu-gothic）")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="入力と配置を検査する")
    mode.add_argument("--catalog", action="store_true", help="レイアウトの名称と用途を表示する")
    args = parser.parse_args(argv)
    # カタログ表示は入力ファイル不要。確認モードは JSON のみで、通常生成は JSON と PPTX 出力先を取る。
    if args.catalog:
        if args.input or args.output:
            parser.error("--catalog にパスは不要です")
        for name, entry in catalog(read_schema()).items():
            print(f"{name:<17} {entry['title']} — {entry['description']}")
        return 0
    if not args.input or (not args.check and not args.output):
        parser.error("入力JSONと出力PPTXを指定してください。--check には入力JSONのみ必要です")
    if args.check and args.output:
        parser.error("--check に出力パスは不要です")
    if args.output and args.output.suffix.lower() != ".pptx":
        parser.error("出力の拡張子は .pptx を指定してください")
    try:
        # utf-8-sig は BOM（ファイル先頭の目印）付き・なしの両方に対応する。
        # 画像の相対パスは実行場所にかかわらず、入力 JSON の親ディレクトリを基準にする。
        document = json.loads(args.input.read_text(encoding="utf-8-sig"))
        slides = render(document, args.input.resolve().parent, page_size=args.page_size, font=args.font)
        # 検査モードでも配置を実行する。文字や表の収まりを確かめ、ファイルへの書き出しを省く。
        if args.check:
            print(f"検査完了。入力 {len(document['slides'])} 件、出力予定 {len(slides)} 枚")
        else:
            write_pptx(args.output, slides, NAVY, ACCENT)
            print(f"{args.output}（{len(slides)} 枚）")
    # 入出力、JSON 解析、配置のエラーを CLI 向けのメッセージと終了コード 1 に変える。
    except (OSError, ValueError) as error:
        parser.exit(1, f"エラー  {error}\n")
    return 0


# ========================================
# プログラムの起動
# ========================================
# このファイルを直接実行した場合に CLI を起動する。他のファイルから import しても起動しない。
if __name__ == "__main__":
    raise SystemExit(main())
