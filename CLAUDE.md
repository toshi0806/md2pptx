# CLAUDE.md

md2pptx — Markdown と PowerPoint テーマ（thmx / pptx）から発表スライド（pptx）を生成するツール。
配色・フォントはテーマに委ね、内容は行頭マーカー記法の Markdown で記述する。

ユーザー向けの使い方は [README.md](README.md)、設計の詳細は [DESIGN.md](DESIGN.md) を参照。
本ファイルはこのリポジトリで作業する際の運用メモ。

## アーキテクチャ

パイプライン（DESIGN.md §2）:

```
theme.thmx ──[thmx2pptx]──▶ base.pptx ┐
                                       ├─▶ [render] ─▶ out.pptx
input.md ──[parser]──▶ IR(Deck) ───────┘
```

| ファイル | 役割 |
|---|---|
| `pyproject.toml` | パッケージ定義（依存・`md2pptx` コンソールスクリプト = `md2pptx.cli:main`） |
| `md2pptx/cli.py` | CLI エントリポイント。引数処理・全体結線（`main()`／`python3 -m md2pptx`） |
| `md2pptx/thmx2pptx.py` | thmx → base pptx 変換（ステージ0）。`theme/`→`ppt/` 等の OPC 操作 |
| `md2pptx/parser.py` | Markdown → 中間表現（IR）。python-pptx 非依存 |
| `md2pptx/ir.py` | IR データクラス（`Deck`/`Slide`/`Line`/`Table`/`Flow`/`TitleSlide` 等）。外部依存なし |
| `md2pptx/render.py` | IR → pptx 描画（`Renderer` クラス）。描画ヘルパーは手書きの参照スクリプトから移植 |
| `md2pptx/flow.py` | フロー図 DSL のパーサ＋座標レイアウタ。python-pptx 非依存（EMU 計算のみ） |
| `md2pptx/pdf.py` | pptx → PDF 変換（`--pdf`）。変換器の探索と外部プロセス起動。python-pptx 非依存 |
| `md2pptx/watch.py` | 入力の変更監視（`--watch`）。stdlib ポーリング。cli にも python-pptx にも非依存 |
| `md2pptx/workdir.py` | 使い捨て作業ディレクトリの作成と片付け。外部依存なし（render / pdf / thmx2pptx が使う） |

パッケージ内モジュールは相対 import（`from .ir import …`）で結線する。
`md2pptx/` はルート直下の flat レイアウト（`pip install .` / `pipx install .` で
`md2pptx` コマンドを生成）。

`parser.py` と `flow.py` は **python-pptx に依存しない純モジュール**（描画は
render の責務）。`ir.py` がパーサとレンダラの契約。新しい記法を足すときは「parser が
IR を作る／render が IR を描く」の分離を保つ。

描画ヘルパー（`box`/`arrow`/`set_autonum`/`no_bullet`/`fit_body` 等）は、
本ツールの土台になった手書きスクリプト（個人デッキ生成用、リポジトリには含めない）から
`render.py` へ移植したもの。

## コマンド

```bash
# 開発中の実行（インストール不要・リポジトリルートで）
python3 -m md2pptx input.md --theme OfficeTheme.pptx -o out.pptx

# インストール後は md2pptx コマンドで実行
md2pptx example.md --theme OfficeTheme.pptx -o example.pptx

# インストール（依存も自動導入）
pipx install .        # 隔離環境（推奨）
pip install -e .      # editable 開発用

# 型チェック（設定は pyproject.toml の [tool.mypy]。CI と同じもの）
pip install -e ".[dev]"
mypy

# 各モジュールの自己検証は -m で（相対 import のみなので直接実行はできない）
python3 -m md2pptx.parser
```

依存: `python-pptx>=1.0` / `PyYAML>=6`（`pyproject.toml` で宣言。インストール時に自動導入）。
環境は python-pptx 1.0.2 / PyYAML 6で検証。

CI（`.github/workflows/ci.yml`）は3ジョブ。`typecheck` が mypy を1回、`generate` が
**3.11 と 3.14 のマトリクス**で `example.md` の生成まで通し、`pdf` が LibreOffice を入れて
`--pdf` のページ数まで検証する（runner に PowerPoint は無いので **LibreOffice 経路のみ**。
macOS の実 PowerPoint 経路は CI では踏めず、手元での確認になる）。mypy は `pyproject.toml` の
`python_version`（= サポート最古）として解析するので実行処理系は1つで足りるが、
**実行マトリクスは別途必要**。mypy が通ることと実際に動くことは別で、#32では `cli.py` の
future import 欠落を実行側だけが捕まえた。

## 変更の検証（重要）

見た目の正しさは **実 PowerPoint レンダリング**で確認する。
python-pptx で開けるだけでは組版のはみ出し等は分からない。

```bash
# 実 PowerPoint(Parallels VM)で PDF 化する手元のツールで out.pptx を変換 → out.pdf
pdftoppm -png -r 110 -f 3 -l 3 out.pdf /tmp/p # 特定ページを画像化 → Read で目視
# 基準と並べる:
magick montage ref.png md.png -tile 2x1 -geometry +4+4 -background '#888' /tmp/cmp.png
```

- 見た目の最終確認は **実 PowerPoint** で行う（macOS なら `--pdf --pdf-converter powerpoint`
  でも同等）。`--pdf` の LibreOffice 経路はフォント解決差で崩れるため、
  当たり確認には使えても最終確認の代替にはならない。
- 実 PowerPoint 変換に使う手元のツールは、
  出力先を指定できて入力の場所にも制約が無いものを想定。
- 構造の確認（枚数・プレースホルダ・フォントサイズ等）は python-pptx で読む。

`md2pptx` 自身にも `--pdf` がある（Issue #39）。生成後に PDF を作る土台機能で、既定
`auto` は native PowerPoint → LibreOffice の順に**使えるものを選ぶ**（無い物は飛ばすが、
**在る物の失敗は握らない**——Issue #46。落とすと忠実度の違う
PDF を黙って掴むことになる）。**macOS では PowerPoint.app があれば `auto` が
osascript 経由で実 PowerPoint を使う**ので、`md2pptx … --pdf --pdf-converter powerpoint`
の出力はそのまま最終確認に使える（LibreOffice 経路は当たり確認どまり）。
初回はオートメーションの TCC 承認が要る（呼び出し元アプリごとに別管理・README 参照）。
`--pdf-converter` に外部の実 PowerPoint 変換ツールを指定することもできる。
忠実度は保証しない（README 参照）。

macOS 経路は **PowerPoint を目立たせずに使う**（Issue #44）。何度変換しても画面が乱れないので、
検証で繰り返し叩いてよい。仕組みは2つで、どちらも消すと運用が壊れる:

- `activate` を入れず `open -g -j -a` で非表示・非アクティブ起動してから文書を開く。
  ウィンドウが出るのは**利用者が既に PowerPoint を表示して使っている場合だけ**で、
  これは仕様（`-j` は起動の瞬間にしか効かない）。`save … as PDF`
  は隠したアプリを自ら再表示するので、起動後に隠し直す方法では抑えられない（測定済み）。
- 変換は **PowerPoint のサンドボックスコンテナ内**で行う（`~/Library/Containers/com.microsoft.Powerpoint/Data/tmp`）。
  pptx をそこへコピーして変換し、PDF を目的地へ移す。
  未承認の場所を直接開かせるとファイルアクセスの許可ダイアログ待ちで止まるが、
  隠していると利用者にはそれが見えないため（実測:
  未承認フォルダ25秒でタイムアウト／コンテナ内 1.1 秒で成功）。

隠したことで気づけない停止（オートメーション承認など）に備え、30秒で
stderr に案内を出す（前面化は tty のときだけ）。

**待ちの上限は tty かどうかで分ける**（Issue #48）。
tty なら打ち切らない——止まる原因の多くは承認ダイアログのような「人が今すぐ直せるもの」で、
30秒の案内はそれを直してもらう仕掛けだから、上から
kill を被せると自分で用意した解決手段を潰すことになる。
非 tty（cron / CI / エディタ拡張）は180秒で打ち切る。`--pdf-timeout` / `MD2PPTX_PDF_TIMEOUT`
で上書き（`0` は無制限）。打ち切り時は自分で起こした子プロセスだけを
kill し、書きかけの PDF を消す。`convert(..., unattended=True)` はこの
tty からの推測を呼び出し側が明示的に打ち消す入口（上限と前面化の両方に効く）。

**出力 pptx / PDF は必ず作業ディレクトリ経由で
`os.replace` する**（`Renderer.save` / `pdf.convert`）。pptx 側の理由は python-pptx が
`zipfile.ZipFile(path, "w")` で書くこと——開いた瞬間に出力先を切り詰めるので、
直接書くと作り直している間だけ壊れた pptx が見える（実測: 25回連続ビルド中38957サンプル中215回。
差し替え方式では0回）。**ただし失敗時の契約は PDF と逆で、
pptx は前回の出力を残す**（pptx の失敗は終了コード1で終わるので取り違えようがなく、
主成果物を消す理由がない）。固定しているのは `tests/test_pptx_replace.py`。

**使い捨て作業ディレクトリの作成と片付けは `workdir.py` に集約する**（Issue #58）。
全5箇所（`render.save` / `pdf.convert` / LibreOffice の使い捨てプロファイル
/ PowerPoint コンテナ内の staging / `thmx2pptx` の展開先）がここを通る。
`workdir.discard` の規則は2つで、どちらも外すと壊れる:
**①片付けの失敗で処理の成否を変えない**（片付けに入る時点で仕事は終わっている。
投げると成功した実行が失敗になり、本体の例外があればそれを置き換えてしまう）、
**②それでも黙って残さない**（`--watch` では保存のたびに1つずつ溜まり、
出力先のものは利用者の目に触れる）。`tempfile.TemporaryDirectory`
に戻してはいけない——**片付けの失敗で例外を投げる**（既定 `ignore_cleanup_errors=False`）ので
① が崩れる。固定しているのは `tests/test_workdir.py`。

PDF 側は加えて、**「変換前に既存
PDF を消す」実装に戻してはいけない**——PDF ビューアはフォルダを監視していて、
削除を確定するとそのファイルを監視から外す（LaTeX Workshop は 250ms で確定し、
集合が空になるとフォルダごと破棄）。変換には1〜数秒かかるので必ず確定してしまい、
**編集しながらのプレビューが最初のリビルドで死ぬ**（実測: 修正前はリビルド1回あたり約1秒 PDF
が不在／修正後は0）。作業場所がファイルではなくディレクトリなのは、LibreOffice が `--outdir` に
`<入力 basename>.pdf` を書くため（`slide.pptx` → `slide.pdf` では出力 PDF を直接書いてしまう）。
無音失敗の検出（存在＋非空）は削除ではなく「毎回まっさらな別名へ書かせる」ことで担保している。
固定しているのは `tests/test_pdf_replace.py`。

`--watch` は入力・テーマ・画像を stdlib のポーリング（0.25 秒）で見張り、
変わるたびに作り直す（Issue #39）。**消すと壊れるもの**が4つ:

- **`build_once` は `SystemExit` を投げない**（`BuildError` を投げる）。投げると watch
  のループごと死に、直して保存しても誰も作り直さなくなる。一発実行のメッセージ・終了コードは
  `main` が `md2pptx: <理由>` へ整形して復元する（`tests/test_cli_build.py` が固定）。
- **監視対象に自分の出力（pptx / PDF）を入れない**。theme に出力 pptx を指されると「作る
  → 変わった → また作る」の無限ループになる。
- **`rebuilding` / `watching for changes` の2行は
  VS Code 側との契約**（`problemMatcher.background` が走行中の判定に使う）。変えるなら
  `.vscode/tasks.json` と VSCODE.md も一緒に直す。`failed to parse|render <md>: <理由>` も同様に
  problemMatcher が読む（Problems パネルの診断）。
- **SIGTERM を `KeyboardInterrupt` へ寄せるのは watch のときだけ**。既定の即死だと
  `finally` が走らず、エディタの「タスクの終了」で PDF 変換の作業ディレクトリが残る。
  一発実行には入れない。

画像パスの解決規則は `render.resolve_image_path`
に集約してある（描画側と監視側で二重管理しない）。

`.vscode/tasks.json` は **VSCODE.md に載せているものと同じ内容**を置いている（ここでは
`example.md` で実際に動かして確かめるため）。**片方だけ直さないこと。** `.vscode/settings.json` の
`editorAssociations` は LaTeX Workshop の内蔵ビューア（自動リロード対応）へ
PDF を向けるもので、これが無いと保存しても画面が変わらない。

## 規約・設計上の約束

- **色・フォントをハードコードしない**。
  図形のみテーマのアクセント色（`self.A2`/`A6`/`T2`/`GOLD`/`BG`/`TX`）を参照する。
  文字サイズは本文/タイトルスタイルから読む（`_body_font_levels`/`_title_font_size`）。
- **表・図のフォント**は本文標準（lvl1）を基本に、
  収まらなければ下位レベルへ段階縮小（`_fit_font`）。見積もりは保守的（安全係数）に。
- **地の文は標準プレースホルダへ**。表・図のあるスライドでも、
  導入文・結論文は本文プレースホルダに入れ、
  空行スペーサで中央帯を空けてオブジェクトを重ねる（自由配置テキストボックスは使わない）。
- 行頭マーカー（`-`/`1.`/`①`/`(1)`/`→`）の解釈は parser に集約。`→` 行は `kind="plain"` で
  no_bullet、ただし **`→` は本文に残す**。
- 丸数字 `①` は文字を除去して `buAutoNum`（`circleNumDbPlain`）へ変換（番号は
  PowerPoint が採番）。

## 落とし穴

- **Bash の stdout 表示が乱れる**ことがある。Python の検証結果はファイルに書き出して
  `Read` で確認すると確実。
- python-pptx の `text_frame.text = "...\v..."` は `\v`(0x0B) を
  `a:br`（行内改行）に展開する。タイトル内 `<br>` は parser で `\v` に変換している。
  `\n` は段落区切り。
- thmx 由来 base はスライド0枚。pptx テーマは既存スライドを持ちうるので、`Renderer.__init__` で
  `_clear_slides()` して常に0枚から描画する（先頭の空きスライド対策）。
- 各サブプロセス/Bash 間で `/tmp` の状態が保持されないことがある。
  生成→検証は1コマンド内で完結させると安全。
- 継承ジオメトリのプレースホルダは、`left`/`width` だけ設定すると
  `top`/`height` が0に落ちる。`_effective_geom` で4辺を解決してから設定する。

## このリポジトリについて

- 親リポジトリ（latex-ecosystem）の `.gitignore` で `*/` 除外されるため、md2pptx は**独立した
  git リポジトリ**。リモート `origin` は `https://github.com/toshi0806/md2pptx.git`。
- 生成物（`example.pptx`/`*.pdf`/`*-slide.pptx`）と Office ロックファイル（`~$*`）は
  `.gitignore` 済み。
- コミットメッセージは英語。PR ワークフローの規約は親リポジトリの CLAUDE.md に従う。
- **文書の日本語は、描画結果が変わらない位置でだけ折る**（Issue #62）。
  段落内の改行は描画時に空白になるので、日本語の途中で折ると本来無い空白が
  GitHub 上に出る。折ってよいのは「句読点の直後」と「すでに空白がある日英の境目」だけで、
  行内コード・リンク・URL の内側では折らない。安全に折れる場所が無い長文はそのまま1行にする。
  **文字列としては同じまま構造だけ壊れる折り方が3つある**ので、
  目でも文字列比較でも気づけない: 閉じ側の `**` の直前で折る（強調が消える）、
  `#` などの行頭記号を行頭に置く（見出しに化ける）、` ```` ` で ` ``` `
  を包んだ**入れ子フェンス**の開閉を取り違える（コードブロックの中身を地の文として畳む）。
  3つとも実際に踏んだ。
- **和欧間の空白は「ラテン文字の語」との間にだけ入れる**（Issue #62）。
  算用数字は日本語の文中では漢数字の代わりなので、語ではなく和文の一部として詰める——`1 つ`
  ではなく「1つ」、`180 秒` ではなく「180秒」、`レイアウト 3` ではなく「レイアウト3」。
  ただし小数点・英字・記号を伴うトークン（`3.11` / `8pt` / `100%`）は語なので空ける（「Python
  3.11 と 3.14 のマトリクス」）。中黒・全角スラッシュの直後も詰める。判定は「数字の隣に
  `.` か英字が付いているか」だけで機械的に決まり、助数詞の語彙リストは要らない。
