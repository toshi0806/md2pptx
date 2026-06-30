# CLAUDE.md

md2pptx — Markdown と PowerPoint テーマ（thmx / pptx）から発表スライド（pptx）を生成する
ツール。配色・フォントはテーマに委ね、内容は行頭マーカー記法の Markdown で記述する。

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
| `src/md2pptx` | CLI エントリポイント（拡張子なし・実行可能）。引数処理・全体結線 |
| `src/thmx2pptx.py` | thmx → base pptx 変換（ステージ0）。`theme/`→`ppt/` 等の OPC 操作 |
| `src/parser.py` | Markdown → 中間表現（IR）。python-pptx 非依存 |
| `src/ir.py` | IR データクラス（`Deck`/`Slide`/`Line`/`Table`/`Flow`/`TitleSlide` 等）。外部依存なし |
| `src/render.py` | IR → pptx 描画（`Renderer` クラス）。`参照スクリプト` のヘルパーを移植 |
| `src/flow.py` | フロー図 DSL のパーサ＋座標レイアウタ。python-pptx 非依存（EMU 計算のみ） |

`parser.py` と `flow.py` は **python-pptx に依存しない純モジュール**（描画は render の責務）。
`ir.py` がパーサとレンダラの契約。新しい記法を足すときは「parser が IR を作る／render が IR を描く」
の分離を保つ。

参照元: `参照スクリプト` は本ツールの土台になった手書きスクリプト（demo スライド生成）。
描画ヘルパー（`box`/`arrow`/`set_autonum`/`no_bullet`/`fit_body` 等）はここから `render.py` へ移植した。

## コマンド

```bash
# 生成（テーマは .thmx / .pptx 両対応）
./src/md2pptx input.md --theme OfficeTheme.pptx -o out.pptx

# デモ一式
./src/md2pptx example.md --theme OfficeTheme.pptx -o example.pptx

# 基準（手書き版）の再生成
python3 参照スクリプト            # → demo-slide.pptx
./src/md2pptx demo.md --theme OfficeTheme.pptx -o demo.pptx
```

依存: `pip install python-pptx pyyaml`（環境は python-pptx 1.0.2 / PyYAML 6）。

## 変更の検証（重要）

見た目の正しさは **実 PowerPoint レンダリング**で確認する。python-pptx で開けるだけでは
組版のはみ出し等は分からない。

```bash
ppt2pdf out.pptx                              # 実 PowerPoint(Parallels VM)でPDF化
pdftoppm -png -r 110 -f 3 -l 3 out.pdf /tmp/p # 特定ページを画像化 → Read で目視
# 基準と並べる:
magick montage ref.png md.png -tile 2x1 -geometry +4+4 -background '#888' /tmp/cmp.png
```

- `ppt2pdf` は **`/Users/toshi/` 配下のファイルしか変換できない**（Windows パスへ写像するため）。
  `/tmp` は不可。リポジトリ内に出力すること（`.gitignore` の `*.pdf`/`*-slide.pptx` を活用）。
- 構造の確認（枚数・プレースホルダ・フォントサイズ等）は python-pptx で読む。

## 規約・設計上の約束

- **色・フォントをハードコードしない**。図形のみテーマのアクセント色（`self.A2`/`A6`/`T2`/
  `GOLD`/`BG`/`TX`）を参照する。文字サイズは本文/タイトルスタイルから読む
  （`_body_font_levels`/`_title_font_size`）。
- **表・図のフォント**は本文標準（lvl1）を基本に、収まらなければ下位レベルへ段階縮小
  （`_fit_font`）。見積もりは保守的（安全係数）に。
- **地の文は標準プレースホルダへ**。表・図のあるスライドでも、導入文・結論文は本文
  プレースホルダに入れ、空行スペーサで中央帯を空けてオブジェクトを重ねる
  （自由配置テキストボックスは使わない）。
- 行頭マーカー（`-`/`1.`/`①`/`(1)`/`→`）の解釈は parser に集約。`→` 行は `kind="plain"`
  で no_bullet、ただし **`→` は本文に残す**。
- 丸数字 `①` は文字を除去して `buAutoNum`（`circleNumDbPlain`）へ変換（番号は PowerPoint が採番）。

## 落とし穴

- **Bash の stdout 表示が乱れる**ことがある。Python の検証結果はファイルに書き出して
  `Read` で確認すると確実。
- python-pptx の `text_frame.text = "...\v..."` は `\v`(0x0B) を `a:br`（行内改行）に展開する。
  タイトル内 `<br>` は parser で `\v` に変換している。`\n` は段落区切り。
- thmx 由来 base はスライド0枚。pptx テーマは既存スライドを持ちうるので、`Renderer.__init__`
  で `_clear_slides()` して常に0枚から描画する（先頭の空きスライド対策）。
- 各サブプロセス/Bash 間で `/tmp` の状態が保持されないことがある。生成→検証は 1 コマンド内で
  完結させると安全。
- 継承ジオメトリのプレースホルダは、`left`/`width` だけ設定すると `top`/`height` が 0 に落ちる。
  `_effective_geom` で 4 辺を解決してから設定する。

## このリポジトリについて

- 親リポジトリ（latex-ecosystem）の `.gitignore` で `*/` 除外されるため、md2pptx は**独立した
  git リポジトリ**（ローカルのみ、リモート未設定）。
- 生成物（`example.pptx`/`demo.pptx`/`*.pdf`/`*-slide.pptx`）と Office ロックファイル（`~$*`）は
  `.gitignore` 済み。
- コミットメッセージは英語。PR ワークフローの規約は親リポジトリの CLAUDE.md に従う。
