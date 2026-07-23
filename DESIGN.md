# md2pptx 詳細設計

**thmx ファイル（テーマ）と Markdown ファイル（内容）から、発表スライド（pptx）を生成するツール。**

スライドのデザインのおおもとは thmx（PowerPoint テーマ）であり、これを一次ソースとして扱う。
中身の記述は Markdown が担う。pptx は最終成果物であって、手作業で用意するテンプレートではない。
`参照スクリプト` のレンダリング資産（色・フォントをハードコードせずテーマに委ねる描画ヘルパー）を再利用する。

## 1. 目的と方針

- **入力は 2 つ：thmx（テーマ＝デザインのおおもと）＋ Markdown（内容）。出力は pptx。**
- **pptx テンプレートを手で用意しない**。thmx から内部でベース pptx を生成し、それを土台に描画する。
  （従来の「テーマだけ持つ手作り pptx」の役割を thmx が直接担う。）
- **汎用の発表スライド生成ツール**。標準的で書きやすい Markdown を主軸にする。
- **色・フォントは thmx のテーマに委ねる**（`参照スクリプト` の思想を踏襲。スクリプトで色をハードコードしない）。
- **行頭マーカー記法**：行の見た目から段落種別を推測して自動変換する。特殊記号を覚えずに書ける。
- **図は独自 DSL**（` ```flow ` フェンスブロック）で box / arrow フロー図に対応する。
- 既存の手書きスライド（参照スクリプト出力）を thmx＋Markdown から再現できるか検証する。

## 2. 全体アーキテクチャ

```
  theme.thmx                         input.md
            │                           │
            ▼                           ▼
 ┌────────────────────┐      ┌────────────────────────────┐
 │ 0. thmx → base pptx │      │ 1. front matter 分離 (YAML) │
 │   （内部一時生成）   │      └────────────────────────────┘
 └────────────────────┘                  │
            │                ┌────────────────────────────┐
            │                │ 2. parser : Markdown → IR   │
            │                │   - スライド分割            │
            │                │   - 行頭マーカー解釈        │
            │                │   - 表 / flow ブロック抽出  │
            │                └────────────────────────────┘
            │                             │
            ▼                             ▼
        ┌──────────────────────────────────────────────┐
        │ 3. renderer : IR → python-pptx               │
        │   - base pptx を土台に描画                    │
        │   - 参照スクリプト のヘルパーを流用           │
        └──────────────────────────────────────────────┘
                              │
                              ▼
                           out.pptx
```

ステージ 0（thmx→base pptx）が本設計の新しい中核。thmx を一度だけ pptx へ変換し、
そのテーマ・スライドマスター・レイアウト一式を持つ「空の pptx」をメモリ上または一時ファイルに用意する。
ステージ 3 はその base pptx に対して `add_slide` していく（従来は手作りの base pptx を開いていた箇所が、
ステージ 0 の生成物に置き換わる）。

モジュール構成（案）:

| ファイル | 役割 |
|---|---|
| `md2pptx.py` | CLI エントリポイント。引数処理・全体オーケストレーション |
| `thmx2pptx.py` | **thmx → base pptx 変換**（ステージ 0）。本ツールの前提を成立させる要 |
| `parser.py` | Markdown → 中間表現（IR）。行頭マーカー・表・flow を解釈 |
| `ir.py` | 中間表現のデータクラス定義 |
| `render.py` | IR → pptx。`参照スクリプト` のヘルパーをライブラリ化して再利用 |
| `flow.py` | ` ```flow ` DSL のパーサ＋レイアウタ（box/arrow 配置計算） |

`参照スクリプト` のヘルパー（`content_slide`, `box`, `arrow`, `note`, `set_autonum`,
`no_bullet`, `fit_body`, `enum_items`, `add_slide_number` 等）は **`render.py` へ移植**し、
IR を受け取って描画する純粋関数群として整理する。

## 3. thmx → base pptx 変換（ステージ 0）

実テーマ（thmx）で実証済みの手順。thmx は内部的に pptx とほぼ同型で、
**差分はわずか 3 点**のため、純 Python（外部アプリ不要）で変換できる。

### 3.1 thmx の構造（実測）

```
[Content_Types].xml
_rels/.rels                         → theme/presentation.xml を officeDocument として参照
theme/presentation.xml              → sldSz 16:9 / sldMasterIdLst / 空の sldIdLst を持つ（実質 presentation.xml）
theme/_rels/presentation.xml.rels   → slideMaster1 を参照
theme/slideMasters/slideMaster1.xml ＋ _rels（layout1..11, theme1 を参照）
theme/slideLayouts/slideLayout1..11.xml ＋ _rels
theme/theme/theme1.xml              → 配色・フォントスキーム本体
theme/theme/themeManager.xml        → テーマパッケージ固有（pptx では不要）
docProps/thumbnail.jpeg ほか画像
```

### 3.2 pptx との差分（この 3 点だけ直す）

1. **パート配置**：`theme/` 配下 → pptx では `ppt/` 配下（rels は相対パスなので中身は無修正で済む）。
2. **コンテンツタイプ**：presentation が `…presentationml.template.main+xml` → `…presentationml.presentation.main+xml`。
3. **Override 欠落**：thmx の `[Content_Types].xml` には slideMaster / slideLayout の Override が無いので追加する。
   併せて themeManager の Override は削除する。

### 3.3 変換アルゴリズム

```
1. thmx を一時ディレクトリへ展開
2. theme/ ディレクトリを ppt/ へリネーム
3. _rels/.rels の Target を ppt/presentation.xml に書き換え（Type=officeDocument のまま）
4. [Content_Types].xml を修正:
     - themeManager の Override を削除
     - presentation の PartName を /ppt/presentation.xml、ContentType を presentation.main+xml に
     - theme1 の PartName を /ppt/theme/theme1.xml に
     - slideMaster1 と slideLayout1..N の Override を追加
5. ZIP し直して base.pptx を得る（[Content_Types].xml を先頭エントリに）
```

> 実証結果（python-pptx 1.0.2）：生成 pptx は正常に開け、サイズ 13.33×7.5（16:9）、
> レイアウト 11 個を認識。レイアウト1「タイトルとコンテンツ」から add_slide → タイトル・
> 本文（レベル付き箇条書き）設定 → 保存まで成功。**従来の描画処理とそのまま接続できる。**

### 3.4 実装上の注意

- レイアウト数 N は `slideLayouts/` 内のファイル数から動的に数える（テーマ差し替えに追従）。
- 生成物は**一時ファイル**（`tempfile`）に置き、レンダリング後に破棄するのを既定とする。
  `--keep-base out-base.pptx` で残せるようにし、デバッグや手直しに使えるようにする。
- thmx 内の画像（背景グラフィック等）は `ppt/` 配下へ移ったあとも rels がそのまま効くので追加対応不要。
- 将来 thmx 以外（既に pptx のテーマ）も許すため、入力拡張子で分岐：`.thmx`→ステージ0、`.pptx`→そのまま base に。

### 3.5 テーマ入力の両対応（thmx / pptx）

テーマは **拡張子で自動分岐**し、どちらも「base pptx」という同一形に収束させる。レンダラ以降は入力形式を意識しない。

```
テーマ入力 ─┬─ .thmx → ステージ0（thmx→base pptx 変換）─┐
            └─ .pptx → 変換せずそのまま base に          ─┴─→ レンダリング（同一経路）
```

```python
def load_base(theme_path):
    ext = os.path.splitext(theme_path)[1].lower()
    if ext == ".thmx":
        base_path = thmx_to_pptx(theme_path)   # ステージ0（一時ファイル）
    elif ext == ".pptx":
        base_path = theme_path                  # そのまま土台に
    else:
        raise SystemExit(f"未対応のテーマ形式: {ext}（.thmx か .pptx を指定）")
    return Presentation(base_path)
```

使い分け:

| 渡すもの | 使いどころ |
|---|---|
| `.thmx` | デザインのおおもとから毎回生成（テーマ更新が即反映される） |
| `.pptx` | thmx から作った base を手直しして固定運用したい／変換をスキップして高速化したい |

## 4. 中間表現（IR）

パーサとレンダラの契約。Markdown の方言や DSL の詳細をレンダラから隠蔽する。

```python
# ir.py （イメージ）
@dataclass
class Line:
    text: str
    level: int             # 箇条書きの深さ 0,1,2...
    kind: str              # "bullet" | "autonum" | "plain"(=no_bullet)
    num_style: str | None  # autonum 時: "arabicPeriod" | "circleNumDbPlain" | "arabicParenBoth" ...
    num_color: str | None  # 採番記号色のテーマ名（例 "tx1"）

@dataclass
class Table:
    header: list[str]
    rows: list[list[str]]

@dataclass
class Flow:                 # ```flow ブロック由来
    nodes: list              # box / ellipsis
    edges: list              # arrow + ラベル
    caption: str | None

@dataclass
class Slide:
    title: str | None
    layout: int              # 既定 1（タイトルとコンテンツ）
    blocks: list             # Line | Table | Flow を順に保持
    directives: dict         # スライド単位の指示（autofit など）

@dataclass
class Deck:
    meta: dict               # front matter
    title_slide: object | None
    slides: list
```

## 5. Markdown 記法仕様

### 5.1 フロントマター（YAML）

ファイル冒頭の `---` で囲む。**テーマ（thmx）・出力先・タイトルスライド情報**を持つ。

```yaml
---
theme: OfficeTheme.pptx            # ★ デザインのおおもと（.thmx / .pptx）。CLI 引数が優先
output: out.pptx                   # CLI 引数が優先
slide_number: true                 # タイトル以外に番号を付与（既定 true）
default_autofit: true              # 本文があふれる場合に縮小（既定 true）

# タイトルスライド（あれば 1 枚目に生成）
title: |
  md2pptx
  Markdown でつくるスライド
subtitle: ― テーマ駆動のスライド生成 ―
author: md2pptx demo
affiliation:
  - Markdown ＋ PowerPoint テーマ → pptx
  - Python / python-pptx / PyYAML
---
```

- 既知キーは `theme` / `output` / `slide_number` / `default_autofit` / `title` /
  `subtitle` / `author` / `affiliation` の 8 つのみ。**未知キーはタイポとみなしてエラー**
  （ディレクティブと同方針。§5.6）。
- キー名は `theme:`（＝thmx）。`.pptx` を渡した場合はステージ 0 を飛ばして直接 base に使う。
- `title` の改行は段落分け（`参照スクリプト` のタイトル多段組みに対応）。
- `subtitle` は副題段落。
- `subtitle` / `author` / `affiliation` の各行は、先頭に本文と同じ相対サイズトークン
  `{-1}` / `{+1}`（5.8）を置ける。基点は**その要素が本来表示されるサイズ**（`subtitle` は
  タイトル枠内に入るためタイトル×0.8、`author` / `affiliation` は副題プレースホルダの既定
  サイズ）で、そこから delta 段調整する。**YAML では値が `{` で始まるとフローマッピングと
  誤解されるため、トークン付きの行はクォートで囲む**（例: `- "{-1} 2026年7月"`）。

### 5.2 スライド分割と見出し

- `## 見出し` で新しいスライドを開始し、見出しがスライドタイトルになる。
- `---`（水平線）でも明示的に分割可（タイトルなしスライドを作りたい場合）。
- `# 見出し`（H1）はセクション見出しスライド（レイアウト2）に割当。
- `###`〜`######`（H3〜H6）は**未定義でエラー**（行番号付き）。将来スライド内小見出しに
  使う可能性を残すため、黙って H2 扱いにしない。
- タイトル内に `<br>`（`<br/>` 可）を書くと、その位置で**タイトルを改行**できる
  （`\v`＝行内改行に変換。`参照スクリプト` の `\v` 相当）。例:
  `## 行頭マーカー記法：<br>見た目で段落の種別を判定`
- 本文行（箇条書き・連番・矢印行）内の `<br>` も同じ規則で `\v` へ変換する
  （マーカー除去後の本文に対して適用。python-pptx が `\v` を段落内改行
  `<a:br/>` として出力するため、続きの行はぶら下げ位置に揃う）。

### 5.3 本文：行頭マーカー記法

インデント（半角スペース 2 つ＝1 レベル）でネストを表す。行頭の見た目で段落種別を決定する。

| 書き方 | 解釈 | 対応する現行処理 |
|---|---|---|
| `- テキスト` / `* テキスト` | 通常箇条書き（テーマ既定の bullet） | `add_bullets` |
| `1. テキスト`（連番） | 自動採番 `arabicPeriod`（1. 2. 3.） | `set_autonum("arabicPeriod")` |
| `①` `②` … で始まる行 | 自動採番 `circleNumDbPlain`（丸数字）。番号文字は除去 | `set_autonum("circleNumDbPlain")` |
| `(1)` `(2)` … で始まる行 | 自動採番 `arabicParenBoth`（丸括弧 (1) (2)） | `set_autonum` 派生 |
| `→ テキスト` で始まる行 | 行頭記号を消した結論・補足行（`plain`） | `no_bullet` |
| 上記以外のプレーン行 | 直前の文脈に従う（既定は bullet level0） | — |

採番の番号色（貢献スライドの番号を黒 `tx1` にする等）はスライド先頭のディレクティブ（5.6）で上書き。既定はテーマ任せ。

- 丸数字は `①`（U+2460）〜`⑳`（U+2473）を認識。他マーカーと違い直後のスペースは省略可。
- `→` 行は「`→` ＋スペース 1 つ＋本文」へ正規化する（サイズトークンの有無で挙動を変えない）。

各行はマーカー直後・本文直前に相対サイズトークン `{+1}` / `{-2}` を置ける（5.8）。例: `- {+1} 強調`、`① {+1} 大きい採番`、`→ {-1} 小さい結論`。

#### 「見出し＋説明」（`enum_items` 相当）

採番行の直下に通常箇条書きをネストすると、見出し=採番 level0／説明=bullet level1 として描画する。

```markdown
1. parser.py
   - Markdown を中間表現（IR）へ変換 → 後述
2. render.py
   - IR を python-pptx で pptx に描画 → 後述
```

### 5.4 表

Markdown 標準のテーブル記法。1 行目をヘッダとしてアクセント色で着色（現行 `add_table` の挙動）。

```markdown
| 課題 | md2pptx のアプローチ |
|---|---|
| デザインの一貫性 | PowerPoint テーマに委譲 |
| 記述のしやすさ | Markdown の行頭マーカー記法 |
```

- 表とテキストを同一スライドに混在可（導入文＋表＋結論）。
- 列幅は均等が既定。`<!-- @table-widths: 45,55 -->` で比率指定できる（任意）。
- 区切り行のコロンで列ごとの水平寄せを指定できる：`:--:` 中央 / `--:` 右 / `:--` 左
  （コロン無し `---` も左＝既定のまま）。

実装メモ（Phase 2）:

- ヘッダ行（区切り行 `|---|` の上の行）を太字＋アクセント色 `A2`／文字色 `BG` で着色。
- **表を含むスライドは座標スタック配置**にする：本文プレースホルダの矩形を内容領域とし、
  ブロック出現順に「テキスト→表→テキスト…」をテキストボックスと表シェイプで重ならないよう縦に積む
  （各セグメントの高さは行数・データ行数による重み配分）。表を含まないスライドは Phase 1 どおり
  本文プレースホルダへ箇条書きを流す（回帰なし）。
- **地の文（導入文・結論文）は標準の本文プレースホルダへ**配置する（自由配置のテキストボックスは使わない）。
  プレースホルダに「導入文＋空行スペーサ＋結論文」を流し、確保した中央帯に表・図を重ねる
  （`参照スクリプト` の図スライドと同方式）。空行数は本文標準サイズの行高から自動算出。
- **表・図のテキストは本文標準（lvl1）サイズを基本**にし、領域に収まらないと概算判定したら
  本文スタイルの下位レベル（lvl2/lvl3…）の小さいサイズへ段階的に切り替える（`_fit_font`）。
  表はセル折り返しを考慮した総高、図 box はラベル＋副ラベルの行数で判定する。
- 最小レベルでも収まらない見積もりの表は警告を出す（PowerPoint は行を最小行高以上へ
  自動拡張するため黙って帯を超過しうる）。縮小させたくない場合は `@overflow: true`
  （§5.6）で本文標準サイズのまま下方向へはみ出せる。

### 5.5 図 DSL（` ```flow ` ブロック）

box / arrow による横並びフロー図を簡潔に書く独自 DSL。`参照スクリプト` の
`box` / `arrow` / `note` の組合せを宣言的に表現する。

````markdown
```flow
direction: lr            # lr(左→右、既定) / tb(上→下)
[theme.thmx | テーマ]
-変換-> [base.pptx | 土台]
-描画-> [out.pptx | スライド]
-> [… | ]                # "…" 単独は省略記号ノード（box ではなく note）
caption: 配色・フォントはテーマ、内容は Markdown
note(top): テーマと Markdown を入力に pptx を生成
note(bottom): → テーマを差し替えるだけで見た目が一新できる
```
````

文法:

- ノード `[ラベル | サブラベル]`。サブラベル省略時は `[ラベル]`。
- ラベルが `…`（または `...`）だけのノードは「省略記号」として box ではなく中央寄せの note にする。
- エッジ `->` は矢印。`-ラベル->` で矢印上にラベル（例 `-PR->`）。
- `direction:` で並び方向（`lr` / `tb` の 2 値のみ。それ以外はエラー）。`caption:` で
  図下キャプション。`note(top|bottom):` で図の上下に注記。
- 配色は thmx のアクセント色を順番に自動割当（`box` の `tc` 引数に相当）。
  `[名前 |サブ]{accent6}` のように末尾 `{themeColor}` で個別指定も許可
  （色名は `accent1`〜`accent6` / `tx1` / `tx2` / `bg1` / `bg2`。`@autonum-color` と共通）。
- ノード・エッジ・設定行のいずれにも解釈できない記述はタイポとみなしてエラー（黙殺しない）。

レイアウト計算（`flow.py`）は、ノード数と方向からスライド中央帯に等間隔配置し、
ノード間に矢印・ラベルを置く。`参照スクリプト` のレビュー工程スライドのロジックを一般化する。

実装メモ（Phase 3）:

- `flow.py` は python-pptx 非依存の純モジュール：`parse_flow(text)→ir.Flow`（`direction`/
  ノード/エッジ/`caption`/`note_top`/`note_bottom`）と、座標プランを返す
  `plan_flow(flow, left, top, width, height)`（EMU 計算のみ）の 2 段に分離。
- 描画は `render.py` の `box`/`arrow`/`note`（`参照スクリプト` から移植）が担い、
  `render_flow` が `plan_flow` のプランを描く。配色は `T2→A6→A6→GOLD→A2` を順に自動割当、
  `{accent6}` 等で個別上書き。
- フロー図は **Phase 2 の座標スタックに `flow` セグメントとして統合**。導入文(Line)→図→結論文(Line)を
  縦に積める（図セグメントの重みはノード数に応じて確保）。`direction: lr`/`tb` の一列フローに対応
  （分岐・格子は将来拡張）。

### 5.6 スライド単位ディレクティブ（任意の上書き）

行頭マーカーで表しきれない指定は、スライド先頭の HTML コメントで補う（必要時のみ）。

```markdown
## 本稿の貢献
<!-- @autonum-color: tx1 -->   # このスライドの採番記号色を黒に
<!-- @layout: 5 -->            # レイアウト番号の上書き
<!-- @autofit: 90 -->          # 本文を 90% に縮小して焼き込む（fit_body scale）
<!-- @body-size: -1 -->        # このスライドの本文を一律で 1 段小さく（5.8）

1. コンテナ・CI・PRレビューを統合した設計の提示
```

ディレクティブの全一覧（値の解釈は各節）:

| ディレクティブ | 意味 | 詳細 |
|---|---|---|
| `@col` | カラム区切り（値を取らない） | §5.7 |
| `@widths: 55,45` / `@widths: 104` | プレースホルダ幅（%）。値 1 個＝単カラム本文幅、複数＝カラムごと。末尾 `!` で左余白使用可 | §5.7 |
| `@table-widths: 45,55` | 表の列幅比 | §5.4 |
| `@overflow: true` | 表・画像の帯からのはみ出し許可（`true`/`false` のみ） | 下記 |
| `@layout: N` | レイアウト番号の上書き | — |
| `@autofit: 90` | 本文の縮小焼き込み（%） | — |
| `@body-size: -1` | スライド既定の相対サイズ段数 | §5.8 |
| `@autonum-color: tx1` | 採番記号色（テーマ色名） | §5.3 |

**overflow の共通規約**（`@overflow` と画像ブロックの `overflow:` に共通）:

- 「帯（セグメント）に収める」制約を外す．**上端は帯上端に固定**し，はみ出しは
  **下（結論文・罫線側）のみ**．タイトル・導入文には重ねない．既定は無効．
- 表: `_fit_font` の段階縮小を行わず，本文標準（lvl1）サイズのまま
  `_table_height_emu` の見積もり高で描画（§5.4）．
- 画像: `@overflow` はスライド既定として作用し，ブロックの `overflow:` 明示
  （`Image.overflow` が True/False．None＝未指定）が優先する（§5.9）．
  明示サイズのない画像は帯に内接するため no-op．
- 帯に収まるオブジェクトには効果なし（通常配置と同じ）．
- 結論文との重なりは仕様（結論文の帯に意図的に重ねる）．実際に帯下端を超え
  結論文があるとき，およびスライド下端を超えるときは stderr に警告．
- Flow は対象外（`plan_flow` が帯に内接固定．将来拡張）．

- **未知のディレクティブは行番号付きエラー**（画像オプションの未知キーと同方針。タイポの
  黙殺を防ぐ）。v0.7 で改名した旧名 `@ph-widths` / `@body-width`（→ `@widths`）・
  `@col-widths`（→ `@table-widths`）は新名称を案内してエラーにする。
- キー名のハイフンはアンダースコアへ正規化（`@body-size` ＝ `@body_size`）。値の区切り・
  記号は全角（`，` `％` `！`）も受理する。

行頭マーカーで日常的な記述はカバーし、ディレクティブは“逃げ道”として最小限に留める。

### 5.7 2 カラム（「2つのコンテンツ」レイアウト）

スライド内に `<!-- @col -->` を 1 つ置くと、その前後が左右 2 つのコンテンツに分かれ、
自動的に「2つのコンテンツ」レイアウト（テンプレートのレイアウト 3）が選ばれる
（`@layout: 3` を明示する必要はない）。区切りの前が左（プレースホルダ idx1）、後が右（idx2）。

```markdown
## 比較：従来方式 と 本環境

- 従来方式（個別TeX導入）
  - OS・バージョン差で環境差
  - 版管理が煩雑・属人化

<!-- @col -->

- 本環境（コンテナ統合）
  - 同一環境を再現・共有
  - PRで版管理・レビュー
```

- 各カラムの中身は通常スライドと同じ行頭マーカー記法（`-`/`1.`/`①`/`→`）が使える。
- IR では `Slide.columns`（各カラムのブロック列）に保持し、レンダラが idx1/idx2 へ流す。
- 箇条書き・採番だけでなく **表・フロー図・画像**もカラム内に置ける。表・図を含むカラムは、その
  カラムのプレースホルダ矩形（`_effective_geom`）を対象に §5.6 と同じスタック描画
  （`_render_stacked_into`）を行う。地の文と表・図が同居する場合は空行スペーサで中央帯を確保して
  棲み分け、表・図しか無ければカラム全体に配置する。Line のみのカラムは従来どおりプレースホルダへ
  流し込む。
- カラム内テーブルの内部列幅比は `@table-widths` で指定できる（スライド単位の指定で、
  全カラムの表に同じ比率を適用。列数が合わない表は等幅フォールバック）。カラム分割比は
  既定でレイアウト 3 依存（通常 50/50）だが、`@widths` でスライド単位に変更できる。
- `<!-- @col -->` を複数置けば 3 カラム以上も IR 上は表せるが、レイアウト 3 のプレースホルダは
  2 つのため、対応プレースホルダが無いカラムは描画されない（テーマ側のレイアウト依存）。

### 5.8 相対フォントサイズ（テーマ基準の段階調整）

見栄えはテーマに委ねる方針を保ったまま、**特定の行だけ**強調・縮小したいときの逃げ道。
絶対 pt は指定できない（テーマ差し替えで破綻するため）。指定できるのはテーマ既定サイズ
からの**相対段数**のみで、1 段あたり ×1.125 / ÷1.125（≈12.5%）。`_fit_font` が下位レベルへ
段階縮小するのと同じ「テーマのサイズ体系の中を上下する」発想。

| 書き方 | 解釈 |
|---|---|
| `- {+1} テキスト` | その行を 1 段大きく（基点＝その行が level から得るテーマ既定サイズ） |
| `1. {-2} テキスト` | 2 段小さく。採番・丸数字・矢印など全行種で使える |
| `→ {-1} テキスト` | トークンは `→` の後ろに置く。`→` は本文に残る |
| `{0}` | テーマ既定に固定（後述のスライド既定 `@body-size` を無効化） |
| `<!-- @body-size: -1 -->` | スライド既定。本文 Line を一律 1 段調整。**行トークンが優先** |

- 符号は省略可（`{2}` ＝ `{+2}`）。`{+0}` / `{-0}` は `{0}` と同義（テーマ既定に固定）。
  トークンが無い行はスライド既定（無ければテーマ既定）に従う。
- 行 `{0}` と `@body-size: 0` は意味が異なる。行 `{0}` は「スライド既定を無効化してその行を
  テーマ既定へ戻す」用途（スライド既定が非 0 のときに効く）。一方 `@body-size: 0` は
  スライド全体で「変化なし」＝既定なしと同義で、何もしない（`@body-size` 無指定と等価）。
- 実サイズ ＝ `round(base × 1.125**delta)` を **8pt〜96pt** でクランプ（極端な段数でも
  暴走しない）。`base` はその行の `level` に対応する本文スタイルの既定サイズ
  （`_body_font_levels`）。pt 値はコードに持たず、テーマ由来の比だけを持つ。
- インデント（`p.level`）は変えず、段落の既定文字書式（defRPr＝`p.font`）にサイズを設定する。
  run の有無に依存せず、bullet・採番記号も本文と同じサイズになる。`@autofit` の縮小とは
  比例関係が保たれるため両立する（相対関係は崩れない）。
- IR では `Line.size_delta`（`int | None`。`None`＝未指定）に保持し、render が実サイズへ換算する。
- front matter の `subtitle` / `author` / `affiliation` でも同じトークンを使える。副題は
  `TitleSlide.subtitle_delta`、著者は `author_delta`、所属は各行対応の `affiliation_deltas`
  に保持する（`affiliation_deltas` は `__post_init__` で `affiliation` と同長に正規化）。
  基点は要素ごとに異なる：`subtitle` はタイトル枠内に入るためタイトル×0.8、`author` /
  `affiliation` は副題プレースホルダ（idx 1）の既定サイズ（`_subtitle_font_size`）。
  YAML の都合でトークン付きの行はクォートが要る（5.1 参照）。

### 5.9 画像（jpg / png）

画像を **表・フロー図と同じ「オブジェクト」**として中央帯にスタック配置する（地の文は
本文プレースホルダへ）。記法は 2 形式（同じ `ir.Image` を生成）。

**(A) ショートハンド（標準 Markdown 画像＋末尾 `{opts}`）**

```markdown
![実験結果の比較](results.png){width=70%}
![](diagram.jpg){width=8cm align=left}
```

- `![キャプション](パス)` の alt を図下キャプションに採用（`opts` の `caption` があれば優先）。
- `opts` は**空白区切り**の `key=value`（`crop` の値はカンマ区切りなので空白では割らない）。
  この制約上、**ショートハンドの `caption=` には空白を含められない**。空白を含むキャプションは
  alt（`![実験 結果](...)`）かフェンス記法（`caption:` 行）を使う。

**(B) フェンス ` ```image `（オプションが多い場合。` ```flow ` と対称）**

````markdown
```image
src: results.png
width: 70%
crop: 100,50,800,400
align: center
caption: 実験結果の比較
```
````

オプション:

| キー | 意味 | 値・既定 |
|---|---|---|
| `src` | 画像（jpg / png） | Markdown ファイルからの相対パス。必須 |
| `width` | 幅 | `70%`（セグメント比）/ `8cm` / `300pt` / `2in` / `150px`。省略時アスペクト自動 |
| `height` | 高 | 同上。省略時アスペクト自動 |
| `crop` | トリミング（残す矩形 `x,y,w,h`） | 既定ピクセル、各値に `%` を付けると割合（全 or 無）。左上原点 |
| `align` | 水平寄せ | `left` / `center`（既定）/ `right` |
| `fit` | width/height 両指定時 | `contain`（既定・比維持で内接）/ `fill`（歪ませ充填） |
| `caption` | 図下キャプション | 省略可 |
| `overflow` | 帯からのはみ出し許可 | `true` / `false`。未指定はスライドの `@overflow`（既定 false）に従い，指定時はそちらより優先。true でクランプ省略・下方向のみはみ出し（共通規約は §5.6） |

- **サイズ**：`width` のみ→高さは比で自動、`height` のみ→幅を自動、両省略→セグメントに内接、
  両指定→既定は内接（比維持）で `fit=fill` のみ歪ませる。最終的にセグメントを超えないよう
  比維持でクランプ（はみ出し防止。表と同じ「帯に収める」方針）。単位無しの数値は px 扱い。
  `overflow: true` はこの最終クランプを外し、明示サイズのまま描画する（`y = max(y, top)` で
  上端はセグメント上端に留め、はみ出しは下方向＝結論文・罫線側のみ。caption も画像下端に追従）。
  `width`/`height` とも未指定の overflow は内接計算のままで意味を持たないため parser で
  エラーにする（`_validate_image`）。描画結果がスライド下端を超えるときは stderr に警告。
- **crop 換算**：ソース画素寸法 `W×H` を `pptx.parts.image.Image.from_file` で読み（Pillow 不要）、
  keep-rect を PowerPoint のクロップ割合へ換算：`cl=x/W, ct=y/H, cr=(W-(x+w))/W, cb=(H-(y+h))/H`
  （`%` 指定は `W=H=100` とみなす）。範囲外は明確なエラーで停止。
- **IR**：`Image(src, width, height, crop, align, fit, caption, overflow)`。`width`/`height` は `Length`
  （`percent` / `emu`）、`crop` は `Crop`（`px` / `percent` の keep-rect）。parser は絶対単位を
  EMU 整数へ換算（`flow.py` と同じ係数）、`%`・crop は render 時に解決（crop はソース寸法が要る）。
- **相対パス**は Markdown ファイルの置き場（`base_dir`、cli が結線）を基準に解決。未検出は fail fast。
- カラム（`@col`）内にも画像を置ける（表・フロー図と同じ経路。§5.7 参照）。

### 5.10 発表者ノート（` ```note ` ブロック）

スライドごとの発表者ノート（オーラル原稿など）を ` ```note ` フェンスで書ける。
スライド面には一切描画せず、pptx の notes slide（発表者ビュー・ノート印刷で見える
テキスト）になる。

````markdown
## 背景：LaTeX論文執筆が抱える課題

- 学生側の障壁
  - OS差・TeX環境・パッケージ依存で環境構築が困難

```note
まず背景です．LaTeX の執筆環境は学生ごとの OS 差で環境構築が難しく，
指導のフィードバックも属人化しがちでした．
```
````

- **宛先**：現在のスライド（直前のスライド開始マーカー以降）。スライド内のどこに
  書いてもよい（慣例は末尾）。1 スライドに複数書いた場合は改行で連結する。
- **タイトルスライド**：本文開始前（最初の見出しより前）に置いた ` ```note ` は
  タイトルスライドのノートになる。フロントマターに `title:` が無い状態で本文開始前に
  置くとエラー（宛先がないため）。
- **段落**：ブロック内の改行は notes の段落区切り（`\n`）としてそのまま保持する。
- **互換性**：`note` を解さない旧版 md2pptx はフェンスを無視するだけなので、
  ノート入りの Markdown は旧版でもビルドできる（ノートが付かないだけ）。
- **IR**：`Slide.notes` / `TitleSlide.notes`（`str | None`）。render は
  `notes_slide.notes_text_frame.text` へ代入するだけ（`_set_notes`）。空なら
  notes slide 自体を生成しない。
- 中身は自由テキスト。行頭マーカー・ディレクティブ等の解釈はしない（コードブロック
  内は素通しという通常のフェンス規則と同じ）。`notes` も別名として受理する。

## 6. レンダリング設計（render.py）

- ステージ 0 の base pptx を開き、`SW/SH`・レイアウト・テーマ色エイリアス（`A2/A6/T2/GOLD/BG/TX`）を初期化。
- IR の各 `Slide` を走査し、`blocks` の型に応じて描画:
  - `Line` 列 → `content_slide` 系（`add_bullets` + マーカーに応じた `set_autonum`/`no_bullet`）。
  - `Table` → `add_table`（ヘッダ着色）。テキストと共存する場合は本文に導入・結論、表は座標配置。
  - `Flow` → `flow.py` のレイアウタで `box`/`arrow`/`note` を配置。
- `default_autofit` が真なら本文プレースホルダに `fit_body`。`@autofit:` 指定があれば scale 焼き込み。
- タイトル以外のスライドに `add_slide_number`。
- テキスト・フォント・色は **テーマ任せ**（図形のみアクセント色を参照）。

## 7. CLI

```bash
pip install python-pptx pyyaml
./md2pptx input.md --theme OfficeTheme.pptx -o out.pptx
```

- 位置引数：Markdown ファイル。
- `--theme`：テーマファイル。**`.thmx` / `.pptx` 両対応**（拡張子で自動分岐。§3.5）。フロントマター `theme:` を上書き。
- `-o/--output`：出力 pptx。フロントマター `output:` を上書き。
- `--keep-base PATH`：ステージ 0 で作った base pptx を破棄せず保存（デバッグ用）。
- 終了時に `saved: <out> slides: <n>` を出力。
- thmx 変換・パースのエラーは「原因＋（パース時は行番号）」を表示して失敗させる。

## 8. 実装フェーズ

1. **Phase 0（thmx 変換）**：`thmx2pptx.py` を実装し、base pptx を確実に生成（実証済みロジックの製品化）。【完了】
2. **Phase 1（最小実用）**：front matter＋タイトルスライド＋`##` 分割＋箇条書き（`-`/`1.`/`→`）＋
   自動採番＋`no_bullet`＋autofit＋スライド番号。`render.py` へヘルパー移植。【完了】
3. **Phase 2（表）**：Markdown テーブル対応、導入文＋表＋結論の混在スライド（座標スタック配置）。【完了】
4. **Phase 3（図 DSL）**：` ```flow ` パーサとレイアウタ、矢印ラベル・caption・note。【完了】
5. **Phase 4（再現検証）**：実在の手書きスライドを Markdown 化し、参照スクリプト出力と
   見比べて差分を詰める（丸数字採番・autofit 率・図の配置など）。【完了】

## 9. 再現検証の結果【Phase 4 完了】

参照スクリプト出力（実在の 24 枚デッキ）を Markdown で再現し，
`ppt2pdf`（実 PowerPoint レンダリング）＋ `pdftoppm` で 1 枚ずつ突き合わせた．
※ 検証に使った個人デッキ・テーマは本リポジトリには含めない。

| 機能 | 想定 Markdown | 再現 |
|---|---|---|
| 多段タイトル＋著者複数行 | front matter | ◎ |
| bullet 多レベル | `-` ネスト | ◎ |
| 導入文＋表＋結論 | 段落＋表＋`→` | ◎ |
| 丸数字採番（色 tx1） | `①`＋`@autonum-color: tx1` | ◎ |
| 一部行のみ丸数字採番（黒） | `①` 混在＋`@autonum-color: tx1` | ◎ |
| enum_items（見出し＋説明） | `1.`＋ネスト `-` | ◎ |
| 全行採番＋結論 no_bullet | `1.` 連番＋`→` | ◎ |
| 本文縮小 | `@autofit: 90` | ◎ |
| box/arrow/note 図 | ` ```flow ` | ◎（矢印は box 高に比例した太さ） |
| 結論行 no_bullet | `→`（記号なし行） | ◎ |
| 多段タイトルの明示改行 | `<br>` | ◎ |

残差（既知）:

- **副題のダッシュ字形**：基準は副題のみゴシック体を明示し「―」を長い全角バーで描くが，
  md2pptx はタイトル枠のフォントを継承するためダッシュがやや短く描かれる（内容は同一）。

タイトルスライドの副題位置・著者枠幅、(2/3) 系タイトルの改行位置は対応済み
（副題はタイトル枠内に少し小さめ＋著者枠を右へ拡張、タイトルは `<br>` で明示改行）。
上記の副題ダッシュ字形以外は，表・採番・enum・チェーン・図（ブロック矢印）・autofit・
no_bullet・タイトル改行を含め実 PowerPoint 上でほぼ同一に再現できることを確認した
（`→` は本文に保持，丸数字は `buAutoNum` 変換）。

## 10. 未決事項 / 留意点

- thmx 変換は実証済みだが、**テーマ差し替え時の堅牢性**（レイアウト数の違い・画像参照・特殊フォント埋め込み）は
  実テーマで都度確認する。Phase 0 でレイアウト数を動的に扱う実装にする。
- 丸数字採番（`①`）は「文字そのまま」ではなく `buAutoNum`（`circleNumDbPlain`）へ変換し番号文字を除去（現行同様）。
- 1 スライドに収まらない量の本文は autofit に頼る（Phase 1）。あふれ警告の要否は後で判断。
- `title` 内のゴシック明示（等幅化回避）など、現行の細かなフォント調整の再現可否は Phase 4 で確認。
- flow DSL の表現力は「横／縦一列のフロー」までを Phase 3 の範囲とし、分岐・格子は将来拡張。
- thmx が「フルテーマ型（マスター＋レイアウト同梱）」であることが前提。色・フォントのみの簡易テーマが
  来た場合はレイアウトが不足するため、その検出と警告を Phase 0 で行う。
