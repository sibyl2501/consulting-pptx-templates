# 閲覧用 PowerPoint 生成キット

調査報告、政策検討、技術解説向けの PowerPoint を、内容を記述した JSON から生成します。既定のA4横（297 × 210 mm）で15種類のレイアウトを使い、結論見出し、図表、考察、出典を1ページで追えるように配置します。

[15種類の見本をダウンロード](./consulting_layout_examples.pptx) · [見本の入力JSON](./examples.json) · [入力仕様](./slide.schema.json)

## 使い方

Python 3.10 以降が必要です。追加パッケージのインストールは不要です。リポジトリのフォルダをターミナルで開き、以下を実行します。

~~~sh
python3 build_slides.py --catalog
python3 build_slides.py --check examples.json
python3 build_slides.py examples.json output.pptx
~~~

従来の横長16:9にする場合は、`python3 build_slides.py --page-size 16:9 examples.json output.pptx` と指定します。

既定の書体はヒラギノ角ゴシック（Hiragino Sans）です。表示環境にない場合などに游ゴシックへ切り替えるには、 `python3 build_slides.py --font yu-gothic examples.json output.pptx` と指定します。

1. `examples.json` をコピーして、資料名と必要なスライドを残します。表紙と目次は必要な資料で追加します。
2. 各スライドの `layout` に型を指定し、見出し、本文、図表の値、出典を書きます。画像のパスは入力JSONの位置から指定します。
3. `--check` で入力と文字の収まりの概算を確認し、PPTXを書き出します。エラーには対象の入力ページと箇所が表示されます。
4. PowerPointで、図表と文字の収まりを確認します。見出し、本文、表、図形を直接編集できます。グラフの数値をまとめて更新する場合は、JSONを変更して再生成します。

### AIに依頼するとき

AIにこのリポジトリのURL、資料の目的・読み手、元資料を渡します。型の選び方も含めて任せるときは、次のように依頼できます。

> `slide.schema.json` と `examples.json` を読み、各ページで何を伝えるかに応じて `layout` を選んでください。見出しには伝えたい結論を置き、本文・図表・出典は元資料に基づいてJSONに記入してください。JSONを `--check` で検査し、エラーのあったページを直してからPPTXを生成してください。

入力JSONを保存しておくと、後で本文や図表を差し替えられます。移動・追加・表の改ページでページ順が変わった場合は、目次のページ番号も更新します。

## レイアウト一覧

| 見本 | `layout` | 使う場面 |
| --- | --- | --- |
| 01 | `cover` | 資料名、対象、範囲を示す |
| 02 | `contents` | 章の読みどころとページを示す |
| 03 | `section` | 章の入口で論点を示す |
| 04 | `summary` | 判断、根拠、次の行動をまとめる |
| 05 | `background` | 背景と今回答える論点を示す |
| 06 | `method` | 調査・実験の条件と手順を示す |
| 07 | `chart` | 棒・折れ線グラフと考察を並べる |
| 08 | `comparison` | 複数案を同じ評価軸で比較する |
| 09 | `data_table` | 詳細な数値や調査結果を掲載する |
| 10 | `annotated_image` | 写真・図面の注目箇所を説明する |
| 11 | `case` | 事例の概要、結果、適用条件を示す |
| 12 | `structure` | 機器・組織・情報の関係を示す |
| 13 | `roadmap` | 担当、実施時期、到達点を示す |
| 14 | `issues` | 課題ごとの影響、対策、担当を対応づける |
| 15 | `mixed` | 全幅と左右2列を組み合わせ、本文・図表・数式を同じページで説明する |

## 入力の要点

- **グラフ**は1点なら右に考察、2点なら下に比較考察を配置します。各系列と横軸のラベル数を揃えます。縦軸を揃えたい場合は両グラフで `min` と `max` を指定します。棒グラフの縦軸は0を含めます。グラフは線・図形・文字で構成し、スライド内で各要素を編集できます。
- **表**には `headers` と `rows` を指定します。`widths` に列幅の比率、`aligns` に左右・中央揃え、`highlight` に強調列の番号を指定できます。`data_table` では、行を分割せずに見出し付きの続きページを作ります。
- **図の注釈**は PNG と JPEG に対応します。`x` と `y` は画像の左上を0、右下を1とする比率です。`alt` は画像の代替テキストとして保存されます。
- **構造図**の `col` と `row`、ロードマップの `start` と `end`、強調列の `highlight` はすべて0始まりです。`links` に同じ `from` を持つ複数の接続を書けば分岐になり、同じ `to` なら合流になります。離れた列間の矢印は下の余白を通ります。同じ列で離れた行同士もつなげます。接続が多い図では、線の交差をPowerPointでも確認してください。
- **文字の収まり**は文字幅と行数を概算して検査します。短い最終行、文字サイズの縮小、文字・表・数式などの重なりの可能性は、検査時と生成時にページ番号と要素名を付けて表示します。基準サイズ（16:9）の本文は標準12 ptで、A4横ではページの横幅に合わせて換算します。フォントや禁則処理によって実際の折り返しが変わるため、最終表示はPowerPointで確認してください。

### 文章・図表・数式を1ページに配置する

`mixed` の `rows` は上から順に並ぶ段です。各段の `blocks` は全幅で描画されます。`right_blocks` を指定するとその段が左右2列になり、各列の要素が上から詰まります。`left_ratio` の既定値は 0.5 で、0.3〜0.7 に指定できます。次の段は左右のうち高い列の下から始まります。本文・表の文字サイズを少し縮めても収まらない場合は、入力ページを示して生成を止めます。

| 要素の `type` | 主な設定 |
| --- | --- |
| `text` | `text` に本文、任意で `heading` に小見出し |
| `equation` | `latex` に数式 |
| `table` | `headers` と `rows`、任意で列幅比の `widths` |
| `chart` | `chart` に棒・折れ線の設定、任意で `height` |
| `image` | `image` に相対パス、`alt` に画像の説明、任意で高さの上限 `height` |

数式はPowerPoint上で編集できます。対応する記法には `x_i^2`、`\frac{a}{b}`、`\sum_{i=1}^{n} x_i`、`\begin{bmatrix}a&b\\c&d\end{bmatrix}`、`\alpha`、`\partial` などがあります。対応外の命令は入力エラーになります。JSONに書くときはバックスラッシュを重ねます。

~~~json
{
  "layout": "mixed",
  "title": "測定値の平均を求める",
  "rows": [
    {
      "blocks": [
        {"type": "text", "heading": "平均", "text": "測定値の合計を回数で割る。"},
        {"type": "equation", "latex": "\\frac{1}{n}\\sum_{i=1}^{n}x_i"}
      ],
      "right_blocks": [
        {"type": "table", "headers": ["回", "測定値"], "rows": [[1, 2], [2, 4]]}
      ]
    }
  ]
}
~~~

## コードを変更する場所

| ファイル | 役割 |
| --- | --- |
| `build_slides.py` | コマンドの入口。入力を検査し、各ページの描画関数を呼ぶ |
| `layouts.py` | 各レイアウト、配色、文字サイズ、図表の配置 |
| `pptx_document.py` | テキスト・図形・表・画像を編集可能なPowerPointとしてパッケージ化 |
| `equations.py` | 対応する LaTeX 記法を PowerPoint の数式構造に変換 |
| `slide.schema.json` | 型ごとの必須項目、入力範囲、用途の説明 |
| `examples.json`、`site_plan.png` | 15種類の見本を作るための入力 |

配色や余白を変える場合は `layouts.py` 冒頭の定数を調整します。型を追加する場合は、同じ識別子を `slide.schema.json`、`layouts.py` の `RENDERERS`、`examples.json` に登録します。

## 参考にした構成

カーニーの要約と本編・付録を横断する目次、BCGの事例分析、デロイト1の指標と施策の対応、デロイト2の業界構造と論点の整理、EYの調査条件と結果を対応づける構成を参考にしました。見本の施設、数値、図面は架空です。

PowerPoint書き出しの土台は、[従来の生成キット](https://github.com/sibyl2501/chat-code-archive/tree/main/projects/consulting-slide-kit) の `pptx_document.py` を流用しました。
