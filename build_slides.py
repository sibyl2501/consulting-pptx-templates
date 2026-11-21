#!/usr/bin/env python3
"""閲覧用のスライドを JSON から生成する。Python 3.10 以降・標準ライブラリのみ。"""
import argparse
import json
import math
from pathlib import Path
import re

from layouts import ACCENT, NAVY, RENDERERS, footer, standard_slide, table_pages
from pptx_document import Slide, write_pptx


SCHEMA_PATH = Path(__file__).with_name("slide.schema.json")
XML_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


class InputError(ValueError):
    """資料の入力箇所が分かるエラー。"""


def read_schema():
    return json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))


def catalog(schema):
    layout_refs = schema["properties"]["slides"]["items"]["oneOf"]
    definitions = schema["$defs"]
    formats = [definitions[item["$ref"].split("/")[-1]] for item in layout_refs]
    return {entry["properties"]["layout"]["const"]: entry for entry in formats}


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
    if "$ref" in rule:
        rule = definitions[rule["$ref"].split("/")[-1]]
    kind = rule.get("type")
    if kind:
        expected = kind if isinstance(kind, list) else [kind]
        if not any(matches_type(value, name) for name in expected):
            raise InputError(f"{at} — {' / '.join(expected)} を指定してください")
    if "const" in rule and value != rule["const"]:
        raise InputError(f"{at} — {rule['const']} を指定してください")
    if "enum" in rule and value not in rule["enum"]:
        raise InputError(f"{at} — 次から選択してください: {', '.join(rule['enum'])}")
    if isinstance(value, str) and ("minLength" in rule and not value.strip()
                                  or XML_CONTROL.search(value)):
        raise InputError(f"{at} — 空白・使用できない制御文字の指定を確認してください")
    for name, predicate, label in (
        ("minimum", lambda limit: value < limit, "以上"),
        ("maximum", lambda limit: value > limit, "以下"),
        ("exclusiveMinimum", lambda limit: value <= limit, "より大きく"),
    ):
        if name in rule and predicate(rule[name]):
            raise InputError(f"{at} — {rule[name]} {label}にしてください")
    if kind == "array":
        count = len(value)
        if count < rule.get("minItems", 0):
            raise InputError(f"{at} — {rule['minItems']} 件以上を指定してください")
        if count > rule.get("maxItems", count):
            raise InputError(f"{at} — {rule['maxItems']} 件以内にしてください")
        for i, item in enumerate(value, 1):
            validate(item, rule["items"], f"{at}[{i}]", definitions)
    if kind == "object":
        props = rule["properties"]
        missing = [name for name in rule["required"] if name not in value]
        if missing:
            raise InputError(f"{at} — 必須項目が不足しています: {', '.join(missing)}")
        for key, item in value.items():
            if key not in props:
                raise InputError(f"{at}.{key} — この項目は定義されていません")
            validate(item, props[key], f"{at}.{key}", definitions)


def validate_document(document, schema):
    kinds = catalog(schema)
    if kinds.keys() != RENDERERS.keys():
        raise RuntimeError("スキーマと描画関数のレイアウト一覧をそろえてください")
    # レイアウト名で分岐を確定し、oneOf の長いエラー一覧を1ページの指摘に絞る。
    header = dict(schema, properties={
        "project": schema["properties"]["project"],
        "slides": {"type": "array", "minItems": 1, "items": {}},
    })
    validate(document, header, "資料", schema["$defs"])
    for i, page in enumerate(document["slides"], 1):
        at = f"入力 {i} 枚目"
        if not isinstance(page, dict):
            raise InputError(f"{at} — object を指定してください")
        name = page.get("layout")
        if not isinstance(name, str) or name not in kinds:
            raise InputError(f"{at}.layout — --catalog にあるレイアウト名を指定してください")
        validate(page, kinds[name], f"{at}（{name}）", schema["$defs"])


def render(document, asset_root, schema=None):
    """書き出しと入力検査で同じ配置を実行する。"""
    schema = read_schema() if schema is None else schema
    validate_document(document, schema)
    slides = []
    for i, page in enumerate(document["slides"], 1):
        kind = page["layout"]
        try:
            records = table_pages(page) if kind == "data_table" else [page]
            for part in records:
                number = len(slides) + 1
                if kind in ("cover", "section"):
                    slide = Slide(f"{number:02d} {part['title']}")
                    if kind == "section":
                        footer(slide, part, number)
                else:
                    slide = standard_slide(part, document, number)
                RENDERERS[kind](slide, part, asset_root)
                slides.append(slide)
        except ValueError as error:
            raise InputError(f"入力 {i} 枚目（{kind}、{page['title']}） — {error}") from error
    return slides


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, nargs="?", help="入力するJSON")
    parser.add_argument("output", type=Path, nargs="?", help="生成するPPTX")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="入力と配置を検査する")
    mode.add_argument("--catalog", action="store_true", help="レイアウトの名称と用途を表示する")
    args = parser.parse_args(argv)
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
        document = json.loads(args.input.read_text(encoding="utf-8-sig"))
        slides = render(document, args.input.resolve().parent)
        if args.check:
            print(f"検査完了。入力 {len(document['slides'])} 件、出力予定 {len(slides)} 枚")
        else:
            write_pptx(args.output, slides, NAVY, ACCENT)
            print(f"{args.output}（{len(slides)} 枚）")
    except (OSError, ValueError) as error:
        parser.exit(1, f"エラー  {error}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
