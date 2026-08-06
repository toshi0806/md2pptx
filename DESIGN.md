# md2pptx 詳細設計

**thmx ファイル（テーマ）と Markdown ファイル（内容）から、
発表スライド（pptx）を生成するツール。**

スライドのデザインのおおもとは thmx（PowerPoint テーマ）であり、これを一次ソースとして扱う。
中身の記述は Markdown が担う。pptx は最終成果物であって、手作業で用意するテンプレートではない。
`参照スクリプト` のレンダリング資産（色・フォントをハードコードせずテーマに委ねる描画ヘルパー）を再利用する。

## 1. 目的と方針

- **入力は2つ：テーマ（デザインのおおもと）＋ Markdown（内容）。出力は pptx。**
- **テーマは手で育てる成果物**とし、その形式は **pptx を第一とする**。
  見た目を変える作業はスライドマスターの編集であり、pptx なら「開く
  → 直す → 保存」で完結する（手順は [THEME.md](THEME.md)）。**thmx
  は変換して取り込む入口**として同格に受け付ける（組織で配布されたテーマをそのまま使える）が、
  編集の往復には書き出しが挟まるので薦めない。
  - 当初は「pptx テンプレートを手で用意しない／thmx を一次ソースとする」方針だった。
    テーマの編集を文書化した時点で**手で保守するテーマが一級市民**になったため、
    上の形へ改めた（Issue #64）。thmx を受け付ける実装は変えていない。
- **汎用の発表スライド生成ツール**。標準的で書きやすい Markdown を主軸にする。
- **色・フォントはテーマに委ねる**（`参照スクリプト` の思想を踏襲。
  スクリプトで色をハードコードしない）。
- **行頭マーカー記法**：行の見た目から段落種別を推測して自動変換する。
  特殊記号を覚えずに書ける。
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

ステージ0（thmx→base pptx）が本設計の新しい中核。thmx を一度だけ
pptx へ変換し、そのテーマ・スライドマスター・レイアウト一式を持つ「空の
pptx」をメモリ上または一時ファイルに用意する。ステージ3はその base pptx に対して
`add_slide` していく（従来は手作りの base pptx を開いていた箇所が、
ステージ0の生成物に置き換わる）。

モジュール構成（案）:

| ファイル | 役割 |
|---|---|
| `md2pptx/cli.py` | CLI エントリポイント。引数処理・全体オーケストレーション |
| `md2pptx/thmx2pptx.py` | **thmx → base pptx 変換**（ステージ0）。本ツールの前提を成立させる要 |
| `md2pptx/parser.py` | Markdown → 中間表現（IR）。行頭マーカー・表・flow を解釈 |
| `md2pptx/ir.py` | 中間表現のデータクラス定義 |
| `md2pptx/render.py` | IR → pptx。`参照スクリプト` のヘルパーをライブラリ化して再利用 |
| `md2pptx/flow.py` | ` ```flow ` DSL のパーサ＋レイアウタ（box/arrow 配置計算） |
| `md2pptx/pdf.py` | pptx → PDF 変換（`--pdf`）。変換器の探索とプロセス起動（python-pptx 非依存） |

パッケージはルート直下の flat レイアウト（`md2pptx/`）で、`pyproject.toml` の
`[project.scripts]` が `md2pptx = "md2pptx.cli:main"` としてコンソールスクリプトを生成する。
モジュール間は相対 import（`from .ir import …`）で結線する。

`参照スクリプト` のヘルパー（`box`, `note`, `set_autonum`, `no_bullet`, `fit_body`,
`add_slide_number` 等）は **`render.py` へ移植**し、IR を受け取って描画する純粋関数群
として整理する。移植したうち IR 経由の描画では通らなくなった 4 つ
（`arrow` / `add_bullets` / `enum_items` / `content_slide`）は削除した（Issue #75）。

## 3. thmx → base pptx 変換（ステージ0）

実テーマ（thmx）で実証済みの手順。thmx は内部的に pptx とほぼ同型で、
**差分はわずか3点**のため、純 Python（外部アプリ不要）で変換できる。

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

### 3.2 pptx との差分（この3点だけ直す）

1. **パート配置**：`theme/` 配下 → pptx では
   `ppt/` 配下（rels は相対パスなので中身は無修正で済む）。
2. **コンテンツタイプ**：presentation が `…presentationml.template.main+xml` →
   `…presentationml.presentation.main+xml`。
3. **Override 欠落**：thmx の `[Content_Types].xml` には slideMaster / slideLayout の
   Override が無いので追加する。併せて themeManager の Override は削除する。

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
> レイアウト11個を認識。レイアウト1「タイトルとコンテンツ」から add_slide → タイトル・
> 本文（レベル付き箇条書き）設定 → 保存まで成功。**従来の描画処理とそのまま接続できる。**

### 3.4 実装上の注意

- レイアウト数 N は `slideLayouts/` 内のファイル数から動的に数える（テーマ差し替えに追従）。
- 生成物は**一時ファイル**（`tempfile`）に置き、レンダリング後に破棄するのを既定とする。
  `--keep-base out-base.pptx` で残せるようにし、デバッグや手直しに使えるようにする。
- thmx 内の画像（背景グラフィック等）は `ppt/` 配下へ移ったあとも
  rels がそのまま効くので追加対応不要。
- 将来 thmx 以外（既に pptx のテーマ）も許すため、入力拡張子で分岐：`.thmx`→ステージ0、
  `.pptx`→そのまま base に。

### 3.5 テーマ入力の両対応（thmx / pptx）

テーマは **拡張子で自動分岐**し、どちらも「base pptx」という同一形に収束させる。
レンダラ以降は入力形式を意識しない。

```
テーマ入力 ─┬─ .thmx → ステージ0（thmx→base pptx 変換）─┐
            └─ .pptx → 変換せずそのまま base に          ─┴─→ レンダリング（同一経路）
```

```python
def load_base(theme_path, keep_base=None):
    """テーマを base pptx の「パス」へ収束させる（Presentation は開かない）．"""
    ext = os.path.splitext(theme_path)[1].lower()
    if ext == ".thmx":
        if keep_base:
            return thmx_to_pptx(theme_path, keep_base), False  # 破棄しない
        return thmx_to_pptx(theme_path), True                  # 一時ファイル
    if ext == ".pptx":
        return theme_path, False                               # そのまま土台に
    raise SystemExit(f"unsupported theme format: {ext}（.thmx か .pptx を指定）")
```

戻り値は `(base_path, is_temp)`。`is_temp` が真のとき、
呼び出し側（`cli.main`）がレンダリング後に一時ファイルを削除する。base pptx を開くのは
`Renderer` の責務で、`load_base` はパス解決だけを行う。

使い分け:

| 渡すもの | 使いどころ |
|---|---|
| `.thmx` | デザインのおおもとから毎回生成（テーマ更新が即反映される） |
| `.pptx` | thmx から作った base を手直しして固定運用したい／変換をスキップして高速化したい |

## 4. 中間表現（IR）

パーサとレンダラの契約。Markdown の方言や DSL の詳細をレンダラから隠蔽する。

以下は骨格のスケッチ。各フィールドの詳細な意味は §5（記法仕様）と `ir.py` の
docstring を正とする。

```python
# ir.py （スケッチ：型と既定値のみ。詳細は ir.py の docstring 参照）
# 既定値のないフィールドは必須（field は dataclasses.field）
Align = Literal["left", "center", "right"]   # 表の列・画像で共通の水平寄せ
TITLE_LAYOUT = 0                             # 表紙レイアウト。parser と render が共有

@dataclass
class Line:                  # 本文の 1 段落
    text: str
    level: int = 0           # 箇条書きの深さ 0,1,2...
    kind: Literal["bullet", "autonum", "plain"] = "bullet"  # plain=no_bullet
    num_style: str | None = None  # autonum 時: "arabicPeriod" | "circleNumDbPlain" | ...
    num_color: str | None = None  # 採番記号色のテーマ名（例 "tx1"）
    size_delta: int | None = None # {+N}/{-N} の相対段数（None=未指定, 0=テーマ既定へ固定）
    # text を "\v" で割ったセグメントと同じ長さ。[0] は常に None（先頭は size_delta）
    seg_deltas: list[int | None] = field(default_factory=list)

@dataclass
class Table:
    header: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    # 列ごとの寄せ（区切り行のコロン由来）。空＝すべて左
    aligns: list[Align] = field(default_factory=list)

@dataclass
class FlowNode:              # フロー図のノード
    label: str = ""
    sublabel: str | None = None
    kind: Literal["box", "ellipsis"] = "box"  # ellipsis は "…" 単独（入力の [...]）
    color: str | None = None # テーマ色名の個別指定（例 "accent6"）

@dataclass
class FlowEdge:              # ノード間の矢印
    src: int = 0             # Flow.nodes 内の index
    dst: int = 0
    label: str | None = None # -PR-> の "PR"

@dataclass
class Flow:                  # ```flow ブロック由来
    direction: Literal["lr", "tb"] = "lr"      # 左→右 / 上→下
    nodes: list[FlowNode] = field(default_factory=list)
    edges: list[FlowEdge] = field(default_factory=list)
    caption: str | None = None
    note_top: str | None = None
    note_bottom: str | None = None

@dataclass
class Length:                # 画像の width / height 1 次元
    unit: Literal["percent", "emu"]  # percent=帯に対する割合 / emu=絶対
    value: float

@dataclass
class Crop:                  # 画像トリミングの「残す矩形」（左上原点）
    unit: Literal["px", "percent"]
    x: float
    y: float
    w: float
    h: float

@dataclass
class Image:                 # ![](){opts} / ```image 由来
    src: str
    width: Length | None = None
    height: Length | None = None
    crop: Crop | None = None
    align: Align = "center"
    fit: Literal["contain", "fill"] = "contain"  # width/height 両指定時
    caption: str | None = None
    overflow: bool | None = None  # None=スライドの @overflow に従う

# 帯へ座標配置するブロック（地の文と違い本文プレースホルダへは入らない）
ObjectBlock = Table | Flow | Image
# スライド本文を構成するブロック（平坦化され Line | Table | Flow | Image と同一）
Block = Line | ObjectBlock

@dataclass
class Slide:
    title: str | None = None
    # title を "\v" で割ったセグメントと同じ長さ。Line と違い [0] も有効
    title_deltas: list[int | None] = field(default_factory=list)
    layout: int = 1          # 既定 1（タイトルとコンテンツ）。表紙は TITLE_LAYOUT
    # 出現順に保持（単一カラム時）
    blocks: list[Block] = field(default_factory=list)
    # スライド単位の指示（autofit など）。値はキーごとに int/str/bool
    directives: dict[str, object] = field(default_factory=dict)
    # 多カラム時の各カラムのブロック列（空なら単一カラム）
    columns: list[list[Block]] = field(default_factory=list)
    notes: str | None = None # ```note 由来の発表者ノート

@dataclass
class TitleSlide:            # front matter 由来（あれば 1 枚目）
    title: str | None = None
    subtitle: str | None = None
    author: str | None = None
    affiliation: list[str] = field(default_factory=list)
    subtitle_delta: int | None = None      # 相対サイズ段数（Line.size_delta と同義）
    author_delta: int | None = None
    # affiliation と 1 対 1（__post_init__ で長さを揃える）
    affiliation_deltas: list[int | None] = field(default_factory=list)
    notes: str | None = None

@dataclass
class Deck:
    meta: dict[str, object] = field(default_factory=dict)  # front matter（YAML 生）
    title_slide: TitleSlide | None = None
    slides: list[Slide] = field(default_factory=list)
```

## 5. Markdown 記法仕様

### 5.1 フロントマター（YAML）

ファイル冒頭の `---` で囲む。**テーマ（thmx）と出力先**を持つ。
表紙は本文記法で書く（§5.2）。

```yaml
---
theme: OfficeTheme.pptx            # ★ デザインのおおもと（.thmx / .pptx）。CLI 引数が優先
output: out.pptx                   # CLI 引数が優先
slide_number: true                 # タイトル以外に番号を付与（既定 true）
default_autofit: true              # 本文があふれる場合に縮小（既定 true）

---
```

- 既知キーは `theme` / `output` / `slide_number` / `default_autofit` と、
  **非推奨**の `title` / `subtitle` / `author` / `affiliation`。
  **未知キーはタイポとみなしてエラー** （ディレクティブと同方針。§5.6）。
- キー名は `theme:`（＝thmx）。`.pptx` を渡した場合はステージ0を飛ばして直接 base に使う。

#### 表紙4項目の非推奨（Issue #82）

`title` / `subtitle` / `author` / `affiliation` は「文書のメタデータ」の名前を持ちながら、
実体は表紙スライドの描画記述だった。裏取りした根拠は3つ。

- **メタデータとして一度も使われていない。** `core_properties` / `docProps` への
  書き込みはコード全体に存在せず、唯一の読み手は `render_title_slide`
- **名前が実体と合っていない。** `affiliation` の例には所属でない行が並ぶ。
  実体は副題プレースホルダ（idx 1）へ入れる行のリスト
- **値の中に描画指示が入る。** `<br>`（折り位置）も `{±n}`（段数）も
  レンダリングの関心事で、その副作用として
  「YAML では `{` で始まる値をクォートせよ」という本文には無い制約が生まれていた

記法が二重にあるので処理も二重になり、`<br>` を片方だけ通し忘れる取りこぼし（#79）まで起きた。

**受理はやめない。** 警告（`_warn_deprecated_meta`）だけ出して従来どおり描く。
非推奨の間は既存原稿が同じ見た目で動き続けるほうが価値がある
（副題の基点を凍結したのと同じ理由。§5.8）。

- `title` の改行は段落分け（`参照スクリプト` のタイトル多段組みに対応）。
- `subtitle` は副題段落。
- `title` / `subtitle` / `author` / `affiliation` の `<br>` は
  行内改行（`\v`）へ変換する（§5.2 と同じ規則）。
  タイトルは枠幅で自動縮小しないため、枠に入らないタイトルの折り位置を
  著者が決める手段でもある。
- `subtitle` / `author` / `affiliation` の各行は、
  先頭に本文と同じ相対サイズトークン `{-1}` / `{+1}`（5.8）を置ける。
  基点は `author` / `affiliation` が副題プレースホルダ（idx 1）の実効既定サイズ、
  `subtitle` だけはマスターの表題サイズ×0.8（実際に描かれるタイトルサイズではない。
  Issue #83 で凍結。理由は 5.8）。**YAML では値が `{` で始まるとフローマッピングと誤解されるため、
  トークン付きの行はクォートで囲む**（例: `- "{-1} 2026年7月"`）。

### 5.2 スライド分割と見出し

- `## 見出し` で新しいスライドを開始し、見出しがスライドタイトルになる。
- `---`（水平線）でも分割できるが**非推奨**（Issue #92）。作られるのは「白紙」ではなく
  「タイトルとコンテンツ」のタイトル枠を空にしたスライドで、テーマが用意していない形。
  加えて CommonMark では段落直後の `---` は setext 見出し（前の行を H2 にする）で、
  行単位で見ている md2pptx とは解釈がずれる。front matter の区切りとも同じ記号。
  受理はやめず、行番号付きの警告だけ出す（front matter の表紙記述と同じ扱い）。
  埋めていた穴は「地の文があってタイトルが無いスライド」だけで、図・表は座標配置なので
  タイトル枠の無いレイアウト（`@layout: 6` の白紙など）にも置ける。
- `# 見出し`（H1）はセクション見出しスライド（レイアウト2）に割当。
- **表紙は `# 見出し` ＋ `<!-- @title-slide -->`**（Issue #82 / #94）。扉との違いは
  レイアウト番号だけで、記法も描画経路も共通。「先頭スライドなら自動的に表紙」という
  位置依存の判定は**採らない**——front matter を排した理由が「書いたものがそのまま出る」
  ことなので、暗黙の位置ルールを持ち込むと同じ問題を別の形で作る。
  主題・副題はタイトル枠、著者・所属は本文行として副題プレースホルダへ入る。
  副題をタイトル枠に置くのはテーマがそう作られているため——多くのテーマは
  タイトル枠と著者欄の間に罫線を引いており、副題を下の枠へ移すと
  罫線の下に落ちて著者情報の塊に混ざる（実 PowerPoint で確認）。
- **`@title-slide` は値を取らず、`@layout` と併記するとエラー**（Issue #94）。
  結果が同じ `@layout: 0` も弾く——「同じ結果なら許す」を入れると、矛盾する
  組み合わせ（`@layout: 5` 等）をどちらで解決するかを別に決めることになる。
  `@layout` は §5.6 で「必要時のみ補う」逃げ道として定義したもので、
  デッキに必ず 1 枚あるものをそれで指定させると、レイアウト番号という
  テーマ側の並びの知識を原稿が持つことになる。レイアウト番号は
  `ir.TITLE_LAYOUT` に置き、parser（`@title-slide` の解決）と
  render（番号を付けない判定・`title_layout`）が同じ値を見る。
- **レイアウト0のスライドには番号を付けない**。そのレイアウトを選ぶこと自体が
  「これは表紙」の宣言なので、番号の有無をそこに紐づける
  （`render_title_slide` が番号を付けないのと揃う）。`slide_number: false` は全体に効く。
- `###`〜`######`（H3〜H6）は**未定義でエラー**（行番号付き）。
  将来スライド内小見出しに使う可能性を残すため、黙って H2 扱いにしない。
- タイトル内に `<br>`（`<br/>` 可）を書くと、
  その位置で**タイトルを改行**できる（`\v`＝行内改行に変換。`参照スクリプト` の
  `\v` 相当）。例: `## 行頭マーカー記法：<br>見た目で段落の種別を判定`
- 本文行（箇条書き・連番・矢印行）内の `<br>` も同じ規則で
  `\v` へ変換する（マーカー除去後の本文に対して適用。python-pptx が `\v` を段落内改行
  `<a:br/>` として出力するため、続きの行はぶら下げ位置に揃う）。
- front matter の `title` / `subtitle` / `author` / `affiliation` も同じ規則。
  ここだけ素通しにすると `<br>` の4文字がそのまま画面に出る（Issue #79）。
  相対サイズトークン `{-1}` の剥がしより**後**に変換する（`_split_size_opt`）。
- `<br>` の直後には相対サイズトークンを置ける（5.8）。変換とトークンの剥がしは
  `_split_br` に集約し、見出し・本文行・矢印行がすべてそこを通る。
  **通す場所を増やしたら全部に掛けること**——#79 は front matter だけ素通しにして
  `<br>` の4文字を画面に出した。

### 5.3 本文：行頭マーカー記法

インデント（半角スペース2つ＝1レベル）でネストを表す。行頭の見た目で段落種別を決定する。

| 書き方 | 解釈 | 対応する現行処理 |
|---|---|---|
| `- テキスト` / `* テキスト` | 通常箇条書き（テーマ既定の bullet） | `_fill_lines`（テーマ既定のまま） |
| `- ` / `-` （マーカーのみ） | 空の段落＝1行空ける | `_mk_bullet`（本文が空でも Line を作る） |
| `1. テキスト`（連番） | 自動採番 `arabicPeriod`（1. 2. 3.） | `set_autonum("arabicPeriod")` |
| `①` `②` … で始まる行 | 自動採番 `circleNumDbPlain`（丸数字）。番号文字は除去 | `set_autonum("circleNumDbPlain")` |
| `(1)` `(2)` … で始まる行 | 自動採番 `arabicParenBoth`（丸括弧 (1) (2)） | `set_autonum` 派生 |
| `→ テキスト` で始まる行 | 行頭記号を消した結論・補足行（`plain`） | `no_bullet` |
| 上記以外のプレーン行 | 直前の文脈に従う（既定は bullet level0） | — |

採番の番号色（貢献スライドの番号を黒 `tx1` にする等）はスライド先頭のディレクティブ（5.6）で上書き。
既定はテーマ任せ。

- 丸数字は `①`（U+2460）〜`⑳`（U+2473）を認識。他マーカーと違い直後のスペースは省略可。
- `→` 行は「`→` ＋スペース1つ＋本文」へ正規化する（サイズトークンの有無で挙動を変えない）。
- **箇条書きマーカーだけの行は空の段落になる**（Issue #82）。記号の出ない枠
  （表紙の著者欄・セクション扉）で行の塊を分けるのに使う。`- {-2}` と書けば
  空行の高さも周りに揃う。**末尾に空白の無い形（`-`）も受ける**——多くのエディタが
  保存時に行末空白を除去するため、`- ` だけを受けると保存した瞬間に壊れる。
  空を段落として残すのは箇条書きマーカーだけで、採番行（`①` / `(1)`）は従来どおり
  行を作らない（番号を1つ消費するだけで、空けたい1行は得られないため）。
  `- <br>` でも空きはできるが**1行多く空く**（段落1行＋`a:br` の2行目）。
  素の空行は従来どおり段落区切りで、段落は作らない。

各行はマーカー直後・本文直前に相対サイズトークン `{+1}` / `{-2}` を置ける（5.8）。
例: `- {+1} 強調`、`① {+1} 大きい採番`、`→ {-1} 小さい結論`。

#### 「見出し＋説明」（採番＋ネスト bullet）

採番行の直下に通常箇条書きをネストすると、見出し=採番
level0／説明=bullet level1 として描画する。

```markdown
1. parser.py
   - Markdown を中間表現（IR）へ変換 → 後述
2. render.py
   - IR を python-pptx で pptx に描画 → 後述
```

### 5.4 表

Markdown 標準のテーブル記法。1行目をヘッダとしてアクセント色で着色（現行
`add_table` の挙動）。

```markdown
| 課題 | md2pptx のアプローチ |
|---|---|
| デザインの一貫性 | PowerPoint テーマに委譲 |
| 記述のしやすさ | Markdown の行頭マーカー記法 |
```

- 表とテキストを同一スライドに混在可（導入文＋表＋結論）。
- 列幅は均等が既定。`<!-- @table-widths: 45,55 -->` で比率指定できる（任意）。
- 区切り行のコロンで列ごとの水平寄せを指定できる：`:--:` 中央
  / `--:` 右 / `:--` 左（コロン無し `---` も左＝既定のまま）。

実装メモ（Phase 2）:

- ヘッダ行（区切り行 `|---|` の上の行）を太字＋アクセント色 `A2`／文字色 `BG` で着色。
- **表を含むスライドは座標スタック配置**にする：本文プレースホルダの矩形を内容領域とし、
  ブロック出現順に「テキスト→表→テキスト…」をテキストボックスと表シェイプで重ならないよう縦に積む（各セグメントの高さは行数・データ行数による重み配分）。
  表を含まないスライドは Phase 1どおり本文プレースホルダへ箇条書きを流す（回帰なし）。
- **地の文（導入文・結論文）は標準の本文プレースホルダへ**配置する（自由配置のテキストボックスは使わない）。
  プレースホルダに「導入文＋空行スペーサ＋結論文」を流し、
  確保した中央帯に表・図を重ねる（`参照スクリプト` の図スライドと同方式）。
  空行数は本文標準サイズの行高から自動算出。
- **表・図のテキストは本文標準（lvl1）サイズを基本**にし、
  領域に収まらないと概算判定したら本文スタイルの下位レベル（lvl2/lvl3…）の小さいサイズへ段階的に切り替える（`_fit_font`）。
  表はセル折り返しを考慮した総高、図 box はラベル＋副ラベルの行数で判定する。
- 最小レベルでも収まらない見積もりの表は警告を出す（PowerPoint
  は行を最小行高以上へ自動拡張するため黙って帯を超過しうる）。縮小させたくない場合は
  `@overflow: true` （§5.6）で本文標準サイズのまま下方向へはみ出せる。

### 5.5 図 DSL（` ```flow ` ブロック）

box / arrow による横並びフロー図を簡潔に書く独自 DSL。`参照スクリプト` の
`box` / 矢印 / `note` の組合せを宣言的に表現する。

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
- ラベルが `…`（または `...`）だけのノードは「省略記号」として box ではなく中央寄せの
  note にする。
- エッジ `->` は矢印。`-ラベル->` で矢印上にラベル（例 `-PR->`）。
- `direction:` で並び方向（`lr` / `tb` の2値のみ。それ以外はエラー）。
  `caption:` で図下キャプション。`note(top|bottom):` で図の上下に注記。
- 配色は thmx のアクセント色を順番に自動割当（`box` の `tc` 引数に相当）。
  `[名前 |サブ]{accent6}` のように末尾 `{themeColor}` で個別指定も許可（色名は
  `accent1`〜`accent6` / `tx1` / `tx2` / `bg1` / `bg2`。`@autonum-color` と共通）。
- ノード・エッジ・設定行のいずれにも解釈できない記述はタイポとみなしてエラー（黙殺しない）。

レイアウト計算（`flow.py`）は、ノード数と方向からスライド中央帯に等間隔配置し、
ノード間に矢印・ラベルを置く。`参照スクリプト`
のレビュー工程スライドのロジックを一般化する。

実装メモ（Phase 3）:

- `flow.py` は python-pptx 非依存の純モジュール：`parse_flow(text)→ir.Flow`（`direction`/
  ノード/エッジ/`caption`/`note_top`/`note_bottom`）と、座標プランを返す
  `plan_flow(flow, left, top, width, height)`（EMU 計算のみ）の2段に分離。
- 描画は `render.py` の `box`/`block_arrow`/`note`（`参照スクリプト` から移植）が担い、
  `render_flow` が `plan_flow` のプランを描く。配色は `T2→A6→A6→GOLD→A2` を順に自動割当、
  `{accent6}` 等で個別上書き。
- フロー図は **Phase 2の座標スタックに `flow` セグメントとして統合**。
  導入文(Line)→図→結論文(Line)を縦に積める（図セグメントの重みはノード数に応じて確保）。
  `direction: lr`/`tb` の一列フローに対応（分岐・格子は将来拡張）。

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
| `@title-slide` | 表紙（値を取らない）。`@layout` との併記はエラー | §5.2 |
| `@widths: 55,45` / `@widths: 104` | プレースホルダ幅（%）。値1個＝単カラム本文幅、複数＝カラムごと。末尾 `!` で左余白使用可 | §5.7 |
| `@table-widths: 45,55` | 表の列幅比 | §5.4 |
| `@overflow: true` | 表・画像の帯からのはみ出し許可（`true`/`false` のみ） | 下記 |
| `@layout: N` | レイアウト番号の上書き（表紙は `@title-slide` を使う） | — |
| `@autofit: 90` | 本文の縮小焼き込み（%） | — |
| `@body-size: -1` | スライド既定の相対サイズ段数 | §5.8 |
| `@autonum-color: tx1` | 採番記号色（テーマ色名） | §5.3 |

**overflow の共通規約**（`@overflow` と画像ブロックの `overflow:` に共通）:

- 「帯（セグメント）に収める」制約を外す．**上端は帯上端に固定**し，はみ出しは
  **下（結論文・罫線側）のみ**．タイトル・導入文には重ねない．既定は無効．
- 表: `_fit_font` の段階縮小を行わず，本文標準（lvl1）サイズのまま
  `_table_height_emu` の見積もり高で描画（§5.4）．
- 画像: `@overflow` はスライド既定として作用し，ブロックの
  `overflow:` 明示（`Image.overflow` が True/False．None＝未指定）が優先する（§5.9）．
  明示サイズのない画像は帯に内接するため no-op．
- 帯に収まるオブジェクトには効果なし（通常配置と同じ）．
- 結論文との重なりは仕様（結論文の帯に意図的に重ねる）．実際に帯下端を超え結論文があるとき，
  およびスライド下端を超えるときは stderr に警告．
- Flow は対象外（`plan_flow` が帯に内接固定．将来拡張）．

- **未知のディレクティブは行番号付きエラー**（画像オプションの未知キーと同方針。
  タイポの黙殺を防ぐ）。v0.7 で改名した旧名 `@ph-widths` / `@body-width`（→
  `@widths`）・`@col-widths`（→ `@table-widths`）は新名称を案内してエラーにする。
- キー名のハイフンはアンダースコアへ正規化（`@body-size` ＝ `@body_size`）。
  値の区切り・記号は全角（`，` `％` `！`）も受理する。

行頭マーカーで日常的な記述はカバーし、ディレクティブは“逃げ道”として最小限に留める。

### 5.7 2カラム（「2 つのコンテンツ」レイアウト）

スライド内に `<!-- @col -->` を1つ置くと、その前後が左右2つのコンテンツに分かれ、
自動的に「2 つのコンテンツ」レイアウト（テンプレートのレイアウト3）が選ばれる（`@layout: 3`
を明示する必要はない）。区切りの前が左（プレースホルダ idx1）、後が右（idx2）。

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
- 箇条書き・採番だけでなく **表・フロー図・画像**もカラム内に置ける。
  表・図を含むカラムは、そのカラムのプレースホルダ矩形（`_effective_geom`）を対象に
  §5.6 と同じスタック描画（`_render_stacked_into`）を行う。
  地の文と表・図が同居する場合は空行スペーサで中央帯を確保して棲み分け、
  表・図しか無ければカラム全体に配置する。
  Line のみのカラムは従来どおりプレースホルダへ流し込む。
- カラム内テーブルの内部列幅比は `@table-widths` で指定できる（スライド単位の指定で、
  全カラムの表に同じ比率を適用。列数が合わない表は等幅フォールバック）。
  カラム分割比は既定でレイアウト3依存（通常50/50）だが、
  `@widths` でスライド単位に変更できる。
- `<!-- @col -->` を複数置けば3カラム以上も IR 上は表せるが、
  レイアウト3のプレースホルダは2つのため、
  対応プレースホルダが無いカラムは描画されない（テーマ側のレイアウト依存）。

### 5.8 相対フォントサイズ（テーマ基準の段階調整）

見栄えはテーマに委ねる方針を保ったまま、**特定の行だけ**強調・縮小したいときの逃げ道。
絶対 pt は指定できない（テーマ差し替えで破綻するため）。
指定できるのはテーマ既定サイズからの**相対段数**のみで、1段あたり ×1.125 / ÷1.125（≈12.5%）。
`_fit_font` が下位レベルへ段階縮小するのと同じ「テーマのサイズ体系の中を上下する」発想。

| 書き方 | 解釈 |
|---|---|
| `- {+1} テキスト` | その行を1段大きく（基点＝その行が level から得るテーマ既定サイズ） |
| `1. {-2} テキスト` | 2段小さく。採番・丸数字・矢印など全行種で使える |
| `→ {-1} テキスト` | トークンは `→` の後ろに置く。`→` は本文に残る |
| `{0}` | テーマ既定に固定（後述のスライド既定 `@body-size` を無効化） |
| `<!-- @body-size: -1 -->` | スライド既定。本文 Line を一律1段調整。**行トークンが優先** |
| `A<br>{-2} B` | `<br>` の直後。その位置から先の run だけ調整（段落は分けない） |
| `# 主題<br>{-2} 副題` | 見出しでも同じ。先頭セグメントにも置ける（`## {+1} X`） |

- 符号は省略可（`{2}` ＝ `{+2}`）。`{+0}` / `{-0}` は `{0}` と同義（テーマ既定に固定）。
  トークンが無い行はスライド既定（無ければテーマ既定）に従う。
- 行 `{0}` と `@body-size: 0` は意味が異なる。行 `{0}`
  は「スライド既定を無効化してその行をテーマ既定へ戻す」用途（スライド既定が非0のときに効く）。
  一方 `@body-size: 0` はスライド全体で「変化なし」＝既定なしと同義で、
  何もしない（`@body-size` 無指定と等価）。
- 実サイズ ＝ `round(base × 1.125**delta)` を **8pt〜96pt** でクランプ（極端な段数でも暴走しない）。
  `base` はその行の `level` に対応する**実効既定サイズ**（`_frame_font_levels`）。
  pt 値はコードに持たず、テーマ由来の比だけを持つ。
- **`<br>` セグメントごとの指定**（Issue #82）。`<br>` は段落を分けずに行を折る記法で、
  python-pptx は `\v` をセグメントごとの `a:r` にして `a:br` でつなぐ。
  つまり run と IR のセグメントが順に1対1で対応するので、run へ書けば
  段落を分けずに一部だけ大きさを変えられる（`_apply_segment_deltas`）。
  タイトル枠の中に副題を収める用途がこれに当たる。
  **段数の基点は行頭の指定から数え直さない**。`- {+1} A<br>{-2} B` の `{-2}` は
  「テーマ既定の2段下」のままで、書く側が位置によらず結果を予測できる。
  格納先は行種で分かれ、理由は書き込む先が違うことにある。
  - 本文行 … 行頭は `Line.size_delta` で**段落の既定文字書式**（`defRPr`）へ。
    そうしないとビュレット・採番記号のサイズが本文とずれる。
    2番目以降は `Line.seg_deltas`（`[0]` は常に `None`）で run へ。
  - タイトル … `Slide.title_deltas` のみ。`[0]` も有効。
    記号が無いので段落側へ書き分ける理由がない。
  - どちらも `__post_init__` でセグメント数と同じ長さに正規化する。
    **長さがずれると別のセグメントにサイズが付く**——見た目の崩れと違い、
    1つずれたサイズは正しく見えてしまうので、構築時に潰す。
- 実効既定サイズは、PowerPoint の継承順（スライド → レイアウト → マスター）に合わせ、
  **レイアウトのプレースホルダの `a:lstStyle` をマスターの `txStyles` より優先**して
  解決する（Issue #83）。テーマは一部のレベルだけ上書きすることがあるのでレベル単位で
  重ねる。マスター側でしか定義されないレイアウト（「タイトルとコンテンツ」など）では
  従来どおりマスターの本文サイズが基点になる。
  マスターのプレースホルダの `a:lstStyle` は見ていない——継承順では間に入るが、
  手元のテーマはどれもそこにサイズを持たず、足しても通らない経路が増えるだけになる。
  なお表・図の段階縮小（`_fit_font`）はプレースホルダではなく図形へ描くため、
  引き続きマスター本文の「サイズの梯子」（`_body_font_levels`）を候補列に使う。
- インデント（`p.level`）は変えず、段落の既定文字書式（defRPr＝`p.font`）にサイズを設定する。
  run の有無に依存せず、bullet・採番記号も本文と同じサイズになる。
  `@autofit` の縮小とは比例関係が保たれるため両立する（相対関係は崩れない）。
- IR では `Line.size_delta`（`int | None`。`None`＝未指定）に保持し、
  render が実サイズへ換算する。
- front matter の `subtitle` / `author` / `affiliation` でも同じトークンを使える。副題は
  `TitleSlide.subtitle_delta`、著者は `author_delta`、所属は各行対応の `affiliation_deltas`
  に保持する（`affiliation_deltas` は `__post_init__` で `affiliation` と同長に正規化）。
  `author` / `affiliation` は副題プレースホルダ（idx 1）へ入るので、基点も本文行と
  同じ `_frame_font_levels` で解決する（同じ枠に出る同じ記法が経路で違うサイズに
  なっていた。Issue #83）。
  `subtitle` だけは基点がマスターの表題サイズ×0.8 のままで、**実際に描かれる
  タイトルサイズ（レイアウトの上書き）を見ていない**。承知のうえで凍結している——
  front matter の表紙記述は Issue #82 で本文記法へ移して非推奨にするため、
  消える直前に見た目を変える益がない。非推奨の間は既存原稿が同じ見た目で動くほうが
  価値がある。#82 で「副題＝タイトル枠に置く行」になった時点で、
  他の行と同じ基点へ合流させる。
  YAML の都合でトークン付きの行はクォートが要る（5.1 参照）。

### 5.9 画像（jpg / png）

画像を **表・フロー図と同じ「オブジェクト」として**中央帯にスタック配置する（地の文は本文プレースホルダへ）。
記法は2形式（同じ `ir.Image` を生成）。

**(A) ショートハンド（標準 Markdown 画像＋末尾 `{opts}`）**

```markdown
![実験結果の比較](results.png){width=70%}
![](diagram.jpg){width=8cm align=left}
```

- `![キャプション](パス)` の alt を図下キャプションに採用（`opts` の
  `caption` があれば優先）。
- `opts` は**空白区切り**の `key=value`（`crop` の値はカンマ区切りなので空白では割らない）。
  この制約上、**ショートハンドの `caption=` には空白を含められない**。
  空白を含むキャプションは alt（`![実験 結果](...)`）かフェンス記法（`caption:` 行）を使う。

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

- **サイズ**：`width` のみ→高さは比で自動、`height` のみ→幅を自動、
  両省略→セグメントに内接、両指定→既定は内接（比維持）で `fit=fill` のみ歪ませる。
  最終的にセグメントを超えないよう比維持でクランプ（はみ出し防止。表と同じ「帯に収める」方針）。
  単位無しの数値は px 扱い。`overflow: true` はこの最終クランプを外し、
  明示サイズのまま描画する（`y = max(y, top)` で上端はセグメント上端に留め、
  はみ出しは下方向＝結論文・罫線側のみ。caption も画像下端に追従）。
  `width`/`height` とも未指定の overflow は内接計算のままで意味を持たないため
  parser でエラーにする（`_validate_image`）。描画結果がスライド下端を超えるときは
  stderr に警告。
- **crop 換算**：ソース画素寸法 `W×H` を
  `pptx.parts.image.Image.from_file` で読み（Pillow 不要）、keep-rect を PowerPoint
  のクロップ割合へ換算：`cl=x/W, ct=y/H, cr=(W-(x+w))/W, cb=(H-(y+h))/H` （`%` 指定は
  `W=H=100` とみなす）。範囲外は明確なエラーで停止。
- **IR**：`Image(src, width, height, crop, align, fit, caption, overflow)`。
  `width`/`height` は `Length` （`percent` / `emu`）、`crop` は `Crop`（`px` / `percent` の
  keep-rect）。parser は絶対単位を EMU 整数へ換算（`flow.py` と同じ係数）、`%`・crop は
  render 時に解決（crop はソース寸法が要る）。
- **相対パス**は Markdown ファイルの置き場（`base_dir`、cli が結線）を基準に解決。未検出は
  fail fast。
- カラム（`@col`）内にも画像を置ける（表・フロー図と同じ経路。§5.7 参照）。

### 5.10 発表者ノート（` ```note ` ブロック）

スライドごとの発表者ノート（オーラル原稿など）を
` ```note ` フェンスで書ける。スライド面には一切描画せず、pptx の
notes slide（発表者ビュー・ノート印刷で見えるテキスト）になる。

````markdown
## 背景：LaTeX論文執筆が抱える課題

- 学生側の障壁
  - OS差・TeX環境・パッケージ依存で環境構築が困難

```note
まず背景です．LaTeX の執筆環境は学生ごとの OS 差で環境構築が難しく，
指導のフィードバックも属人化しがちでした．
```
````

- **宛先**：現在のスライド（直前のスライド開始マーカー以降）。
  スライド内のどこに書いてもよい（慣例は末尾）。1スライドに複数書いた場合は改行で連結する。
- **表紙**：本文記法で書いた表紙（`@title-slide`）は普通のスライドなので、
  ノートもその中に書く（宛先の規則は上と同じ）。
- **非推奨のタイトルスライド**：本文開始前（最初の見出しより前）に置いた
  ` ```note ` はフロントマター由来のタイトルスライドのノートになる。
  `title:` が無い状態で本文開始前に置くとエラー（宛先がないため）。
- **段落**：ブロック内の改行は notes の段落区切り（`\n`）としてそのまま保持する。
- **互換性**：`note` を解さない旧版 md2pptx はフェンスを無視するだけなので、ノート入りの
  Markdown は旧版でもビルドできる（ノートが付かないだけ）。
- **IR**：`Slide.notes` / `TitleSlide.notes`（`str | None`）。render は
  `notes_slide.notes_text_frame.text` へ代入するだけ（`_set_notes`）。空なら
  notes slide 自体を生成しない。
- 中身は自由テキスト。行頭マーカー・ディレクティブ等の解釈はしない（コードブロック内は素通しという通常のフェンス規則と同じ）。
  `notes` も別名として受理する。

## 6. レンダリング設計（render.py）

- ステージ0の base pptx を開き、
  `SW/SH`・レイアウト・テーマ色エイリアス（`A2/A6/T2/GOLD/BG/TX`）を初期化。
- IR の各 `Slide` を走査し、`blocks` の型に応じて描画:
  - `Line` 列 → 本文プレースホルダへ `_fill_lines`／`_append_lines` で流し込み、
    マーカーに応じて `set_autonum`/`no_bullet` を適用（`bullet` はテーマ既定のまま）。
  - `Table` → `add_table`（ヘッダ着色）。テキストと共存する場合は本文に導入・結論、
    表は座標配置。
  - `Flow` → `flow.py` のレイアウタで `box`/`block_arrow`/`note` を配置。
- `default_autofit` が真なら本文プレースホルダに `fit_body`。`@autofit:` 指定があれば
  scale 焼き込み。
- タイトル以外のスライドに `add_slide_number`。
- テキスト・フォント・色は **テーマ任せ**（図形のみアクセント色を参照）。
- **描画した run には言語（`a:rPr/@lang`）を付ける**（`_apply_text_language`。Issue #79）。
  PowerPoint は**行分割の規則を run の言語で選ぶ**。
  python-pptx は `paragraph.text = …` で `<a:r><a:t>…</a:t></a:r>` を作り、
  `a:rPr` を一切書かない。
  何もしないと言語が決まらないまま出力され、**日本語の禁則処理が適用されない**。
  行頭に「ー」や句読点が来るが、文字列としては正しいので pptx を開くまで気づけない。
  - 効くのは `lang` **だけ**。
    `kumimoji="1"` だけでは効かない（縦書き中の数字の扱いであって行分割ではない）。
    `presentation.xml` の `<p:kinsoku>`（禁則文字の定義）を足しても効かない。
    いずれも実 PowerPoint 変換で確認した。
  - **`render` の最後に一度だけ通す**。
    run を作る経路は本文・タイトル・表・フロー図・ノートと多く、
    作る側それぞれに足すと必ずどこかが漏れる。
    `save` ではなく `render` に置くのは、
    言語が「描いた文字の性質」であって「ファイルの書き方」ではないため
    （`save` の責務はアトミック差し替えに限る）。
  - **既に決まっている run は上書きしない**。何度通しても結果が変わらず、
    run 単位で言語を決めるようにしたときも後から潰さない。
    テーマが自分で持ち込む run（`add_slide_number` が複製する番号枠の
    「Page」など）の言語もそのまま残る。
  - 言語は `_LANG` に固定してある。md2pptx は日本語のスライド用なので
    設定項目は増やさない。**拡張が要るならこの定数を置き換えるのではなく、
    run ごとに言語を決める処理を前段に足す**——上書きしない性質があるので、
    先に決めた言語はこの処理を通しても変わらない。
    front matter に `lang:` を足すならその前段として実装する。
  - レイアウト・マスターは触らない（テーマの所有物）。
    ノートも通す（発表者ノートの折り返しも同じ理由で崩れる）。
- **出力 pptx はアトミックに差し替える**（`Renderer.save`。Issue #56）。
  出力先と同じディレクトリに使い捨ての作業ディレクトリ（`.md2pptx-*`）を作ってそこへ保存し、
  `os.replace` で置き換える。
  作業場所は出力先の**中に新しく作る**ので必ず同一ファイルシステム上にあり、
  置き換えは常にアトミック（EXDEV は起こりえない）。PDF（§7の `convert()`）と同じ形。
  - 理由は **python-pptx が `zipfile.ZipFile(path, "w")` で書く**こと。
    開いた瞬間に出力先を切り詰めるので、直接書くと「作り直している間だけ壊れた
    pptx が見える」時間が生まれる。実測（`example.md` を25回連続ビルドしながら
    0.5ms 間隔でサンプリング）で **38957サンプル中215回**、
    zip として開けない状態が観測された（差し替え方式では43968サンプル中0回）。
    PowerPoint で開いたまま作り直す、`--watch` と手で叩くのを併用する、
    といった使い方で普通に踏む。
  - **失敗したら前回の pptx を残す**——ここは PDF と**逆の契約**。
    PDF 変換の失敗は終了コードを変えない（警告だけ）ので、古い
    PDF を残すと「新しい出力」と取り違えられる。pptx の保存失敗は cli が
    `BuildError` にして終了コード1で終えるため取り違えようがなく、
    それなら主成果物（PowerPoint で開いているかもしれない）を消さない方がよい。
  - `--keep-base` の中間 base pptx は対象外（デバッグ用の副産物で、`thmx2pptx` の担当）。
- **使い捨て作業ディレクトリの作成と片付けは `workdir.py` に集約する**（Issue #58）。
  5箇所（`render.save` / `pdf.convert` / LibreOffice の使い捨てプロファイル
  / PowerPoint コンテナ内の staging / `thmx2pptx` の展開先）が同じ規則で動く。
  - **片付けの失敗で処理の成否を変えない。** 片付けに入る時点で保存も変換も終わっている。
    ここで投げると成功した実行が失敗になり、
    本体が投げた例外があればそれを握りつぶして置き換えてしまう。
    `tempfile.TemporaryDirectory` は既定（`ignore_cleanup_errors=False`）で
    **片付けの失敗に例外を投げる**ので使わない——実測で `PermissionError` が送出され、それは
    `PdfError` ではないため cli の `except PdfError` を素通りして終了コード1になる（「PDF
    が作れなくても pptx は成功」§7が崩れる）。
  - **それでも黙って残さない。** 消せなかったと誰も知らないと
    `--watch` では保存のたびに1つずつ溜まる。とくに `render.save` と
    `pdf.convert` の作業場所は**出力先ディレクトリ**に作るので利用者の目に触れる。
    原因は環境側（Windows で走査ソフトがファイルを掴んでいる等）なので再試行はせず、
    起きた事実だけを stderr に出す。
  - **作成**の失敗は整形しない（呼び出し側の事情で違う）。`pdf.convert` は `PdfError`
    にして終了コードを変えず、`render.save` は何をしようとして失敗したかを添えて送出する。

## 7. CLI

```bash
pipx install .                                            # 依存も自動導入（開発中は python3 -m md2pptx）
md2pptx input.md --theme OfficeTheme.pptx -o out.pptx
```

- 位置引数：Markdown ファイル。
- `--theme`：テーマファイル。**`.thmx` / `.pptx` 両対応**（拡張子で自動分岐。§3.5）。
  フロントマター `theme:` を上書き。
- `-o/--output`：出力 pptx。フロントマター `output:` を上書き。
- `--keep-base PATH`：ステージ0で作った base pptx を破棄せず保存（デバッグ用）。
- `--pdf`：pptx 生成後に PDF も作る（値は取らない）。出力先は出力
  pptx と同じ場所・basename の `.pdf`。変換は `pdf.py` が担う。
  **PDF 変換だけ失敗しても終了コードは0**（警告のみ）——編集しながらのプレビューを止めないため。
  忠実度は変換器による（README 参照）。
- `--pdf-output PATH`：PDF の出力先。**単独で指定しても生成を有効にする**（`--keep-base PATH`
  と同じ形。出力先を書いた人が「作るな」を意図することはない）。
  `--pdf` と併用しても矛盾ではないので、その場合は PATH に作る。
  - `--pdf` に値を持たせる形（`nargs="?"`）は採らない。
    `md2pptx --pdf deck.md` で入力ファイルが `--pdf` の値として食われ、
    「input が無い」という原因の分からないエラーになるため（#42）。
  - 有効化するのは**成果物を名指しするオプションだけ**。
    `--pdf-converter` は「どう作るか」の指定なので有効化しない——`MD2PPTX_PDF_CONVERTER` を
    export しただけで全実行が PDF を作り始めてしまう。
- `--pdf-converter NAME|COMMAND`：PDF 変換器。`auto`（既定・PowerPoint→LibreOffice）/
  `powerpoint` / `libreoffice` / 任意コマンド（`{input}`/`{output}`/`{outdir}` 置換）。
  環境変数 `MD2PPTX_PDF_CONVERTER` を上書き。
  - `auto` がするのは**探索だけ**で、失敗の肩代わりはしない（#46）。
    「その変換器が無い」（`_Unavailable`）ときだけ次へ進み、
    **在る物が失敗したらそのまま失敗**させて `--pdf-converter libreoffice` を案内する。
    落としてしまうと忠実度という成果物の性質が黙って入れ替わるうえ、
    隠れる原因（オートメーション承認の拒否・ライセンス未認証など）は利用者が直せるものだから。
    可用性の判定は macOS が app バンドルの有無、Windows は
    COM オブジェクトを作れたか（PowerShell に成功マーカーを出させて切り分ける）、
    LibreOffice は `soffice` が見つかるか。
  - macOS の `powerpoint` は
    `osascript`（`save … in (POSIX file p) as save as PDF`）で実 PowerPoint に変換させる。
    `POSIX file` への coerce は必須（POSIX パス文字列だと保存先を解決できない）。
    オートメーションの TCC 承認は**呼び出し元バイナリごと**に別管理で、
    未承認だと承認待ちで止まる（README の注意参照）。
  - PowerPoint は**非表示で起動し**（`open -g -j -a`・AppleScript に `activate` を入れない）、
    **自分のサンドボックスコンテナの中だけを触らせる**（pptx をそこへコピーして変換し、
    PDF を目的地へ移す）。前者は変換のたびに画面を奪わないため、
    後者はフォルダごとのアクセス承認を要求させないため——隠して動かしている以上、
    承認ダイアログで止まっても利用者には見えないので、そもそも出さない（#44）。
    隠したことで気づけない残りの停止（オートメーション承認など）に備え、
    30秒で stderr に案内を出し PowerPoint を前面に出す（変換は中断しない）。
    ただし**前面化は stderr が tty のときだけ**（#48）——非対話の呼び出し元では、
    そこで出ているのは*呼び出し元アプリ*に対する承認ダイアログなので
    PowerPoint を前面化しても押せず、作業中の画面からフォーカスを奪うだけになる。
  - **待ちの上限**（#48）：止まり方には「人が今すぐ直せるもの」（承認ダイアログ・サインイン画面）と「誰も直さないもの」（GUI
    セッションの無い cron / CI・クラッシュ）があり、時計では区別できない。そこで **stderr が
    tty か**で分ける——tty なら打ち切らない（30秒の案内は人に直してもらうための仕掛けで、
    その上から kill を被せると自分で用意した解決手段を潰す。`Ctrl-C` で止められる）、
    非 tty なら180秒で打ち切る。`--pdf-timeout` / `MD2PPTX_PDF_TIMEOUT`
    で上書き（`0` は無制限）。上限は全変換器に共通で適用し、
    補助コマンド（LaunchServices への問い合わせ）は人が介在しないので10秒固定。
    打ち切り時は**自分で起こした子プロセスだけ**を kill し（PowerPoint 本体は殺さない）、
    書きかけの PDF を消す。
    - `convert(..., unattended=True)` は、この **tty
      からの推測を呼び出し側が明示的に打ち消す** 入口。上限の決め方と前面化の両方に効く。tty
      から読み取っているのは本来「人が今その画面を見ていて直せるか」であって「端末が端末か」ではないので、
      呼び出し側がそれを知っているなら、そちらを優先させる。
  - **出力 PDF はアトミックに差し替える**。
    出力先と同じディレクトリに使い捨ての作業ディレクトリ（`.md2pptx-*`）を作ってその中で変換し、
    成功したときだけ `os.replace` で目的のパスへ移す。理由は3つ。
    - **変換中に出力 PDF が消えている時間を作らない**のが第一。
      PDF ビューアはフォルダを監視していて、
      削除を確定するとそのファイルを監視集合から外す（LaTeX Workshop は既定
      250ms で確定し、集合が空になるとフォルダの監視ごと破棄する。`onDidChange`
      は集合に無いファイルを即 return する）。変換は実 PowerPoint で1〜2秒・LibreOffice
      で数秒かかるので必ず確定してしまい、**以後どれだけ作り直しても再読込されない**。
      実測でも、消してから書く実装ではリビルド1回あたり約1秒 PDF
      が不在になっていた（差し替え方式では0）。
    - `save as PDF` は印刷パイプライン経由で**無音失敗しうる**（exit 0でも
      PDF が無い・空）ので、終了コードを信じず**成果物の存在＋非空**を成功条件にしている。
      以前はそのために変換前へ既存出力の削除を置いていたが、
      **毎回まっさらな別名へ書かせる**方が同じ保証を削除なしで与える（「前回の
      PDF が成功条件を満たしてしまう」状況が構造的に作れない）。
    - ファイルではなく**ディレクトリ**なのは、出力パスを取らない変換器のため。
      LibreOffice は `--outdir` に `<入力 basename>.pdf` を書くので、
      `slide.pptx` → `slide.pdf` という既定の組み合わせでは**変換器が出力
      PDF を直接・逐次的に書いていた**（ビューアが書きかけを読む）。作業ディレクトリを挟むと
      `outdir` がそちらへ移り、この衝突も同時に消える。
    - 失敗したときに古い PDF を残さない契約は維持する（PDF 変換の失敗は終了コードを変えないので、
      警告を見落とした人が前回の内容を新しい出力と取り違えてしまう）。
      書きかけは作業ディレクトリごと捨てる。
- `--watch`：入力とその依存を見張り、変わるたびに作り直し続ける（`watch.py`）。
  - **監視は stdlib のポーリング**（0.25 秒）で行い、`watchdog` のような依存は足さない。
    見張るのは「そのビルドが実際に読んだファイル」＝ Markdown・テーマ・画像の数個で、
    ディレクトリ全体ではないので、1周あたり数個の `os.stat` にしかならない。
    必要なのは「変わったか」だけでイベントの種別も要らず、`sleep`
    を注入できるぶんテストが決定的になる。ディレクトリを丸ごと見張りたくなったら再考する。
  - 指紋は `(st_mtime_ns, st_ctime_ns, st_size, st_ino)`。`mtime` だけでは足りない——解像度が1秒の
    FS で、同じ秒に同じ長さで書き直されると見落とす。**不在は `None`** として持つので、
    「消えた」「置かれた」も変更として拾える（足りない画像を置いたら作り直したい）。
  - **監視対象は毎回ビルドの戻り値で入れ替える**。
    **自分の出力（pptx / PDF）は必ず除く**——theme に出力 pptx を指されると「作る → 変わった
    → また作る」の無限ループになる。
  - **変化を見つけても1周期待ってから作り直す**。
    エディタの保存は書き込みを複数回に分けることがあり、大きい画像のコピーは途中経過が見える。
    指紋は**ビルド前**に取る——PDF 変換の数秒がまるごと死角になり、
    その間の保存を「反映済み」と誤認してしまうため。
    - ただし待つのは `SETTLE_LIMIT`（10秒）まで。
      書き込みが止まらない相手（別プロセスが書き換え続けるファイル、
      非常に遅い回線越しのコピー）だと落ち着く瞬間が来ず、**一度も作り直さないまま黙り込む**。
      無言は失敗の仕方として最悪で、利用者からは watch が壊れたようにしか見えない。
      上限に達したら理由を出して作りに行く——中途半端なファイルで失敗しても
      watch は止まらず、落ち着いた後の変更でまた作り直せる。
  - **失敗しても止まらない**。編集中は失敗しているのが普通の状態で、
    ここで終了すると直して保存しても誰も作り直さない。
    失敗時は前回の依存も残す（足りない画像が置かれたら作り直す）。
    ただし起動時の引数検査（入力の不在・空の `--pdf-output`）はその場で終了する。
  - **SIGTERM を `KeyboardInterrupt` へ寄せる**（watch のときだけ）。既定の SIGTERM は即死で
    `finally` が走らず、エディタの「タスクの終了」で作業ディレクトリが残る。
    `Ctrl-C` と同じ経路に寄せて後片付けを通し、**意図した停止なので終了コードは0**。
  - PDF 変換には `unattended=True` を渡す（180秒で打ち切り／前面化しない）。
    tty から読み取っているのは本来「人が今その画面を見ていて直せるか」で、
    watch では見ているのはエディタと PDF であってタスクの端末ではない。
  - 進捗は stderr に `rebuilding` / `watching for changes` の2行（時刻つき）。
    **この2行はエディタ側の契約**で、VS Code の `problemMatcher.background`
    が走行中かどうかの判定に使う。
  - 速度目的ではない（実測：pptx 側は合計 0.2 秒弱で、うち 0.13 秒がインタプリタ起動。
    支配項は PDF 変換の1〜数秒）。常駐で縮むのは1割に満たないので、
    LibreOffice を listener 常駐させるような複雑さは持ち込まない——必要な人には
    `--pdf-converter 'unoconvert {input} {output}'` という出口が既にある。
  - **入力 Markdown はビルド前から見張り始める**（`run(seed=...)`）。依存が判明するのは
    `build` が返った後なので、これが無いと初回だけは比較の起点を持てず、
    **初回ビルド中の保存が「反映済み」に見えて消える**。実 PowerPoint の初回は実測6秒かかり、
    起動してすぐ書き始めると普通に踏む（手元で実際に踏んだ）。取りこぼしが残るのは「seed
    に無い依存（テーマ・画像）を初回ビルド中に書き換えた場合」だけ。
- `--version`：バージョンを表示して終了（`md2pptx.__version__` が単一の情報源）。
- 終了時に `saved: <out> slides: <n>`（`--pdf` 時はさらに `saved: <pdf>`）を出力。
- thmx 変換・パースのエラーは「原因＋（パース時は行番号）」を表示して失敗させる。
- 1回ぶんのビルドは `build_once`（`BuildError` / `BuildResult`）に切り出してあり、
  一発実行と watch が共有する。**`build_once` は `SystemExit` を投げない**——投げると
  watch のループごと死ぬ。一発実行では `main` が `md2pptx: <理由>` の `SystemExit`
  に整形するので、利用者から見たメッセージと終了コードは切り出し前と同一（唯一の例外が
  `unsupported theme format` で、他と揃って `md2pptx: ` 接頭辞が付くようになった）。

## 8. 実装フェーズ

1. **Phase 0（thmx 変換）**：`thmx2pptx.py` を実装し、
   base pptx を確実に生成（実証済みロジックの製品化）。【完了】
2. **Phase 1（最小実用）**：front matter＋タイトルスライド＋`##`
   分割＋箇条書き（`-`/`1.`/`→`）＋自動採番＋`no_bullet`＋autofit＋スライド番号。
   `render.py` へヘルパー移植。【完了】
3. **Phase 2（表）**：Markdown テーブル対応、導入文＋表＋結論の混在スライド（座標スタック配置）。
   【完了】
4. **Phase 3（図 DSL）**：` ```flow ` パーサとレイアウタ、矢印ラベル・caption・note。
   【完了】
5. **Phase 4（再現検証）**：実在の手書きスライドを Markdown 化し、
   参照スクリプト出力と見比べて差分を詰める（丸数字採番・autofit 率・図の配置など）。
   【完了】

## 9. 再現検証の結果【Phase 4完了】

参照スクリプト出力（実在の24枚デッキ）を Markdown で再現し，
実 PowerPoint レンダリング＋ `pdftoppm` で1枚ずつ突き合わせた．
※ 検証に使った個人デッキ・テーマは本リポジトリには含めない。

| 機能 | 想定 Markdown | 再現 |
|---|---|---|
| 多段タイトル＋著者複数行 | front matter | ◎ |
| bullet 多レベル | `-` ネスト | ◎ |
| 導入文＋表＋結論 | 段落＋表＋`→` | ◎ |
| 丸数字採番（色 tx1） | `①`＋`@autonum-color: tx1` | ◎ |
| 一部行のみ丸数字採番（黒） | `①` 混在＋`@autonum-color: tx1` | ◎ |
| 見出し＋説明 | `1.`＋ネスト `-` | ◎ |
| 全行採番＋結論 no_bullet | `1.` 連番＋`→` | ◎ |
| 本文縮小 | `@autofit: 90` | ◎ |
| box/矢印/note 図 | ` ```flow ` | ◎（矢印は box 高に比例した太さ） |
| 結論行 no_bullet | `→`（記号なし行） | ◎ |
| 多段タイトルの明示改行 | `<br>` | ◎ |

残差（既知）:

- **副題のダッシュ字形**：基準は副題のみゴシック体を明示し「―」を長い全角バーで描くが，
  md2pptx はタイトル枠のフォントを継承するためダッシュがやや短く描かれる（内容は同一）。

タイトルスライドの副題位置・著者枠幅、
(2/3) 系タイトルの改行位置は対応済み（副題はタイトル枠内に少し小さめ＋著者枠を右へ拡張、
タイトルは `<br>` で明示改行）。上記の副題ダッシュ字形以外は，
表・採番・enum・チェーン・図（ブロック矢印）・autofit・no_bullet・タイトル改行を含め実 PowerPoint
上でほぼ同一に再現できることを確認した（`→` は本文に保持，丸数字は `buAutoNum` 変換）。

## 10. 未決事項 / 留意点

- thmx 変換は実証済みだが、
  **テーマ差し替え時の堅牢性**（レイアウト数の違い・画像参照・特殊フォント埋め込み）は実テーマで都度確認する。
  Phase 0でレイアウト数を動的に扱う実装にする。
- 丸数字採番（`①`）は「文字そのまま」ではなく
  `buAutoNum`（`circleNumDbPlain`）へ変換し番号文字を除去（現行同様）。
- 1スライドに収まらない量の本文は autofit に頼る（Phase 1）。あふれ警告の要否は後で判断。
- `title` 内のゴシック明示（等幅化回避）など、現行の細かなフォント調整の再現可否は
  Phase 4で確認。
- flow DSL の表現力は「横／縦一列のフロー」までを Phase 3の範囲とし、分岐・格子は将来拡張。
- thmx が「フルテーマ型（マスター＋レイアウト同梱）」であることが前提。
  色・フォントのみの簡易テーマが来た場合はレイアウトが不足するため、その検出と警告を
  Phase 0で行う。
