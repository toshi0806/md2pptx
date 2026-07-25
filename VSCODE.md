# VS Code で編集しながらプレビューする

LaTeX Workshop のリアルタイムプレビューと同じ体験になります。**Markdown を保存すると数秒で
PDF のタブが更新されます**（タブは開いたまま、スクロール位置も保たれます）。

以下は **macOS を前提**に書いています。Windows / Linux でも同じ仕組みで動きますが、キーは
`Cmd` を `Ctrl` に読み替えてください（オートメーション承認は macOS だけの話です。Windows の
PowerPoint 変換は COM 経由なので承認は要りません）。

- [用意するもの](#用意するもの)
- [タスクを置く](#タスクを置く)
- [使う（通しの手順）](#使う通しの手順)
- [キーバインド](#キーバインド)
- [自動で始める（`runOn`）](#自動で始めるrunon)
- [うまくいかないとき](#うまくいかないとき)

## 用意するもの

**1. [LaTeX Workshop](https://marketplace.visualstudio.com/items?itemName=James-Yu.latex-workshop)**（`James-Yu.latex-workshop`）

`.tex` と無関係の PDF も開けて、**ディスク上の更新を検知して自動で読み直す**ビューアを持って
いるのが理由です。LaTeX を使っていなくても構いません。

- `vscode-pdf`（`tomoki1207.pdf`）は自動リロードが壊れたままなので使えません
  （[Issue #162](https://github.com/tomoki1207/vscode-pdfviewer/issues/162)）。
- VS Code 内にこだわらないなら、外部ビューアでも同じことができます（macOS なら
  [Skim](https://skim-app.sourceforge.io/) が自動リロードに対応。プレビュー.app は非対応）。

**2. PDF をそのビューアに関連付ける**

`.vscode/settings.json`（またはユーザー設定）に:

```jsonc
{
  "workbench.editorAssociations": { "*.pdf": "latex-workshop-pdf-hook" }
}
```

**この設定は形式的なものではありません。** PDF ビューアを持つ拡張を複数入れていると（`vscode-pdf`
など）、どちらが開くかはこの設定で決まります。自動リロードするのは LaTeX Workshop だけです。

## タスクを置く

原稿のリポジトリの `.vscode/tasks.json` に貼ってください（このリポジトリにも同じものが入って
いるので、`example.md` ですぐ試せます）。

```jsonc
{
  "version": "2.0.0",
  "tasks": [
    {
      // 編集中の Markdown を見張り続ける。
      "label": "md2pptx: watch",
      "type": "shell",
      "command": "md2pptx",
      "args": ["${file}", "--watch", "--pdf"],
      "options": { "cwd": "${fileDirname}" },
      "isBackground": true,
      "presentation": {
        // 端末は出さない（PDF を見ていればよい）。問題は Problems パネルに出る。
        // 端末も見たいときは "always" にする。
        "reveal": "silent",
        "panel": "dedicated",
        "clear": true,
        "close": false
      },
      "problemMatcher": {
        "owner": "md2pptx",
        "source": "md2pptx",
        "fileLocation": ["autoDetect", "${fileDirname}"],
        "severity": "error",
        "pattern": {
          // 拾うのはファイルに紐づく失敗だけ（`converting to PDF: ...` を診断にしない）。
          "kind": "file",
          "regexp": "^md2pptx: failed to (?:parse|render) (.+?\\.md): (.+)$",
          "file": 1,
          "message": 2
        },
        "background": {
          "activeOnStart": true,
          "beginsPattern": "^md2pptx: \\d\\d:\\d\\d:\\d\\d rebuilding\\b",
          "endsPattern": "^md2pptx: \\d\\d:\\d\\d:\\d\\d watching for changes\\b"
        }
      }
    },
    {
      // 単発ビルド（Cmd-Shift-B の既定）。
      "label": "md2pptx: build",
      "type": "shell",
      "command": "md2pptx",
      "args": ["${file}", "--pdf"],
      "options": { "cwd": "${fileDirname}" },
      "group": { "kind": "build", "isDefault": true },
      "presentation": { "reveal": "silent", "panel": "dedicated", "clear": true },
      // pattern は上の watch と同じもの。tasks.json にはカスタム matcher を名前で
      // 使い回す仕組みが無いので写すしかない——**直すときは 2 か所とも直すこと**。
      "problemMatcher": {
        "owner": "md2pptx",
        "source": "md2pptx",
        "fileLocation": ["autoDetect", "${fileDirname}"],
        "severity": "error",
        "pattern": {
          "kind": "file",
          "regexp": "^md2pptx: failed to (?:parse|render) (.+?\\.md): (.+)$",
          "file": 1,
          "message": 2
        }
      }
    },
    {
      // 出来上がった PDF を開く（上の関連付けで内蔵ビューアへ）。`code` コマンドが
      // PATH に要る（コマンドパレット → 「Shell Command: Install 'code' command in PATH」）。
      "label": "md2pptx: open preview",
      "type": "shell",
      "command": "code",
      "args": ["-r", "${fileDirname}/${fileBasenameNoExtension}.pdf"],
      "presentation": { "reveal": "never" },
      "problemMatcher": []
    }
  ]
}
```

## 使う（通しの手順）

**0.** リポジトリを VS Code で開く

**1. Markdown を開いて、エディタ内をクリックする**

これが**いちばん間違えやすいところ**です。タスクの `${file}` は**実行した瞬間にアクティブな
エディタ**に固定されます。先に PDF を開いてしまうと `${file}` が PDF になり、Markdown ではなく
PDF を見張ることになります。

**2. タスクを実行する** — どちらでも同じです

- **メニューバー**: 「ターミナル」→「タスクの実行...」（Terminal → Run Task...）
- **コマンドパレット**: `Cmd+Shift+P` → `タスク: タスクの実行`（英語 UI なら `Tasks: Run Task`）

一覧が出るので **`md2pptx: watch`** を選びます。`problemMatcher` を書いてあるので、「タスクの
出力をスキャンせずに続行しますか」は聞かれません。

**3. 動いているか確認する**

`"reveal": "silent"` にしてあるので**端末は出てきません**（PDF を見ていればよい、という設計）。
初回は不安になるので、確かめ方を書いておきます。

`Cmd+J` でパネルを開き、「ターミナル」タブ右のドロップダウンに `md2pptx: watch` があれば動いて
います。そこを選べば進捗が見えます。

```
md2pptx: watching example.md — Ctrl-C to stop
md2pptx: 14:03:21 rebuilding example.md
saved: example.pptx slides: 14
saved: example.pdf
md2pptx: 14:03:23 watching for changes
```

**4. PDF を横に開く**

エクスプローラで PDF を右クリック →「横に開く」（Open to the Side）。

**5. Markdown を編集して保存する**

数秒で右の PDF が更新されます。タブは開いたまま、スクロール位置も保たれます。

**6.（確認）文法エラーを入れて保存してみる**

`Cmd+Shift+M` で**「問題」パネル**を開くと、こう出ます。

```
example.md
  ⊗ image not found: example-fig-typo.png    md2pptx
```

右端の `md2pptx` が `problemMatcher` の `source` です。直して保存すれば消えます。

**7. 止める** — 「ターミナル」→「タスクの終了...」

別の Markdown に切り替えるときも、いったん止めてから手順 1 に戻ってください（`${file}` は
タスク実行時に固定されるため）。原稿が 1 つに決まっているなら、`${file}` の代わりに
`"${workspaceFolder}/slide.md"` と書いておく方が事故がありません。

## キーバインド

`keybindings.json` へ（`args` では変数が展開されないので、タスク名で呼びます）。

```jsonc
[
  { "key": "cmd+alt+w", "command": "workbench.action.tasks.runTask",
    "args": "md2pptx: watch", "when": "editorLangId == markdown" },
  { "key": "cmd+alt+v", "command": "workbench.action.tasks.runTask",
    "args": "md2pptx: open preview", "when": "editorLangId == markdown" },
  { "key": "cmd+alt+t", "command": "workbench.action.tasks.terminate" }
]
```

`workbench.action.tasks.terminate` は **md2pptx 専用ではありません**。実行中のタスクが複数あると
どれを止めるか尋ねられるので、そこで `md2pptx: watch` を選んでください。

## 自動で始める（`runOn`）

`runOptions.runOn` を使うと、**フォルダを開いた時点で watch が始まります**。手順 1〜3 が丸ごと
不要になるので、原稿が 1 つに決まっているならこちらが快適です。

```jsonc
{
  "label": "md2pptx: watch",
  "args": ["${workspaceFolder}/slide.md", "--watch", "--pdf"],
  "runOptions": { "runOn": "folderOpen" },
  // 以下は同じ
}
```

`${file}` ではなく**固定パスにするのが要点**です。フォルダを開いた時点ではまだどのエディタも
開かれていないので、`${file}` は当てになりません。

代償が 2 つあります。**既定にしていないのはこのため**です。

- **フォルダを開くたびに PowerPoint 変換が走ります**（`--pdf` を付けている場合）。ちょっと
  ファイルを見に来ただけのときにも動きます。
- **初回に「このフォルダの自動タスクを許可するか」の確認**が出ます（設定
  `task.allowAutomaticTasks`）。許可するまでは動きません。

## うまくいかないとき

**保存しても PDF が変わらない**

- PDF のタブが LaTeX Workshop のビューアで開かれているか確認してください。エディタ右上が
  PDF 表示ではなくテキストや別のビューアなら、[用意するもの](#用意するもの) の関連付けが
  効いていません。
- タスクが動いているか（手順 3）。`${file}` が PDF に固定されていないか（手順 1）。

**「問題」パネルに md2pptx と関係ないログが出る**

それは**「出力」パネル**かもしれません。パネル下部のタブは **問題 / 出力 / デバッグ コンソール /
ターミナル** と並んでいます。`2026-… [info] …` のような**タイムスタンプ付きの行は「出力」**で、
拡張機能が書いたログです（C# Dev Kit などはワークスペースを開くと言語に関係なく動きます）。

**「問題」パネルにはファイル名・行番号・メッセージしか出ません。** md2pptx の診断は `source` が
`md2pptx` になります。`Cmd+Shift+M` で「問題」タブへ直接行けます。

**テーマ未指定などのエラーが Problems に出ない**

`problemMatcher` が拾うのは**ファイルに紐づく失敗だけ**です（`failed to parse|render <md>: <理由>`）。
`no theme specified` のようにファイルを特定できないエラーは端末にだけ出るので、見落としが
気になるなら `presentation` の `"reveal"` を `"always"` にしてください。

**初回に PowerPoint の承認ダイアログが出る**

macOS のオートメーション承認で、**呼び出し元アプリごと**に別管理です。Terminal で承認済みでも
VS Code では改めて出ます。押し損ねても 180 秒で諦めるだけなので、次の保存でまた出ます。
PowerPoint を使いたくない場合は `--pdf-converter libreoffice` を `args` に足してください
（GUI も承認も不要。ただし忠実度は当たり確認どまりです）。

**`md2pptx: open preview` が PDF を開けない**

このタスクは「Markdown の隣に同じ名前の PDF ができる」前提です（`slide.md` → `slide.pdf`）。
フロントマターの `output:` が別の場所や名前を指しているときは、`args` をその PDF のパスに
書き換えてください。一度開いてしまえば以降はタブが自動更新されるので、使うのは最初の 1 回だけです。

**タスク機構を疑いたい**

切り分けには、統合ターミナルで直接叩くのが手っ取り早いです。

```bash
md2pptx slide.md --watch --pdf
```

これで PDF が更新されるならビューア側は正常で、問題はタスク定義の側にあります。
