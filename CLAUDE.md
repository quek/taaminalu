# taaminalu - TSF対応ターミナルエミュレータ

TSF `ITextStoreACP` を実装し、ターミナルバッファの内容を IME に公開するターミナルエミュレータ。
Rust + Win32 (Direct2D/DirectWrite) で実装。WSL2 に ConPTY 経由で接続。

## 目的

既存のターミナルアプリは TSF テキストストアの読み取りに対応していない（`GetText` が空を返す）。
rtry (Try-Code IME) のストロークヘルプ等、カーソル位置の文字を読み取る IME 機能が動作しない。
このターミナルは IME が `GetText` でバッファ内容を読み取れるようにする。

## アーキテクチャ

```
WSL2 shell/tmux
    ↕ stdin/stdout
wsl.exe (ConPTY 経由で起動)
    ↕ ReadFile/WriteFile
alacritty_terminal (VT パース + グリッドバッファ)
    ↓                ↓
ITextStoreACP     Direct2D/DirectWrite
(TSF テキストストア)   (レンダリング)
```

## 主要コンポーネント

| モジュール | 責務 |
|---|---|
| app | アプリケーション状態（タブ一覧、Renderer、アクティブタブ管理） |
| tab | 個別タブ（PTY + TermWrapper のペア） |
| pty | ConPTY で wsl.exe/cmd.exe/powershell を起動、読み書き |
| term | `alacritty_terminal` のラッパー。グリッドバッファ・ACP 変換 |
| tsf | `ITextStoreACP` + `ITfContextOwnerCompositionSink` 実装。IME にバッファ公開 + composition 管理 |
| render | Direct2D + DirectWrite でターミナル描画（preedit インライン描画含む） |
| window | Win32 ウィンドウ作成・メッセージループ・IMM32 候補位置制御 |

## ビルド・実行

```
cargo build --release
cargo run --release
```

- **GUI アプリ (`#![windows_subsystem = "windows"]`)** のため `eprintln!` の出力は見えない
- ビルド時に `target/release/taaminalu.exe` がロックされている場合はアプリを閉じてから再ビルド

## 開発ルール

### 応答言語
- **日本語で応答する**

### コミット
- コミットメッセージは日本語で書く

### コーディング
- KISS・DRY
- edition 2024、最新の crate バージョン
- コードを書いたら都度レビュー・リファクタリング
- **コンパイル警告（warning）を残さない**。コミット前に全て解消すること
- **要件にない変更を入れるな**。既存の挙動（デフォルト値、初期状態等）を勝手に変えない
- **エラーを握りつぶすな**。`unwrap_or_default()` / `ok()` で安易に無視せず、根本原因を調査してから対処する

### コミット前レビュー（必須）
コミット前に変更箇所を対象に以下の観点でレビューし、問題があれば修正する。

**パフォーマンス:**
- 描画ループ内のヒープ確保（Vec, String, format!）→ スタック配列・キャッシュを使う
- TSF コールバック内のヒープ確保（GetText 等は数十回連続で呼ばれる）→ `Arc<String>` 等で clone を軽量化
- COM オブジェクト（Brush, TextLayout）の毎フレーム生成 → キャッシュする
- Mutex の二重ロック → 1回のロックで必要な値をまとめて取得
- O(n) の繰り返し呼出 → 結果のキャッシュや軽量な代替メソッドを検討
- ファイル I/O を伴うデバッグログ → コミット前に必ず削除（高頻度メソッドでは致命的）

**セキュリティ:**
- `unsafe` ブロック: ポインタの null チェック・境界チェックがあるか
- 整数キャスト (`as i32`, `as u16`, `as usize`): 切り捨て・オーバーフローの可能性 → `saturating_add` 等を使う
- 外部入力（クリップボード、PTY、IME）: バッファサイズの検証があるか
- `from_raw_parts` / `copy_nonoverlapping`: 長さの妥当性を検証しているか
- エラーの握りつぶし (`?` → `unwrap_or_default()` / `ok()`): 根本原因を調査し、原因そのものを修正する

### 調査・デバッグ
- **推測でコードを書くな**。まずコードを読んで原因を追ってから修正する
- **安易な解決策を採用するな**。類似プロダクトの実装やベストプラクティスを調べてから実装する
- **デバッグは上流から下流へ**: 関数が呼ばれているか → 引数は正しいか → ロジックは正しいか
- **前提を検証してから実装する**: 「呼ばれるはず」等の前提はログで検証してから次に進む
- 「可能性がある」ではなく、確実に原因を特定してから修正する
- windows crate の API は `~/.cargo/registry/src/` 内のソースを grep して確認する

#### 描画バグの調査手順
表示がおかしい場合、以下の順序で調査する（上流→下流）:
1. **グリッドダンプを最初にやる**: セルの文字・フラグ・カーソル位置をファイルに出力し、データが正しいか確認する
2. **データが正しい場合**: 描画コード（render.rs）がセルのフラグ（INVERSE, BOLD, HIDDEN 等）を全て処理しているか確認する
3. **データが間違っている場合**: VT パーサー（alacritty_terminal）の処理を確認する
4. VT ストリーム操作やグリッド後処理は**最後の手段**。まずデータと描画の両方を検証してから検討する

教訓: カーソル位置ずれの原因が「文字幅不一致」だと推測し、VT ストリーム操作→ unicode-width パッチ→グリッド後処理と試行錯誤したが、実際の原因は「INVERSE フラグの描画未対応」だった。グリッドダンプを最初にやっていれば、データが正しいことが即座にわかり、描画コードの確認だけで解決できた。

#### COM/TSF バグの調査手順
COM/TSF の問題はコードレビューだけでは原因を特定できない。**必ずログで検証する。**
1. **セットアップの成功を最初に確認する**: `setup_tsf` の各ステップ（`CoCreateInstance` → `Activate` → `CreateDocumentMgr` → `CreateContext` → `AdviseSink` → `Push` → `SetFocus`）が成功しているかログで検証する。`.ok()` / `?` でエラーが握りつぶされて初期化自体が失敗しているケースがある
2. **コールバックが呼ばれているか確認する**: `AdviseSink`, `RequestLock`, `GetText`, `OnStartComposition` 等にログを仕込み、TSF マネージャーから実際に呼ばれているか検証する
3. **ログ出力先は `std::env::temp_dir()`**: GUI アプリは `eprintln!` が見えないため、ファイルに出力する
4. **推測で修正するな**: 「フォーカス管理が原因かも」と推測して修正するのではなく、ログで実際の失敗箇所を特定してから修正する

教訓: GetText が動かない原因を「TSF フォーカス管理の欠如」と推測して `AssociateFocus` を追加したが、実際の原因は `ITfSource::AdviseSink` が `CONNECT_E_ADVISELIMIT` を返し `?` で `setup_tsf` 全体が失敗していたこと。ログで `setup_tsf` の各ステップを検証していれば 1 往復で解決できた。

教訓: rtry 交ぜ書き変換でバックスペース後にテキストが消える問題を、ConPTY のコードページ・バイト分割・遅延送信と推測で何度も試したが、PTY write のバイト列をログで確認したら通常 Backspace との違い（まとめて送信 vs 1バイトずつ）が即座に判明した。**推測での修正は1回まで、2回目以降はログで事実確認。**

#### ConPTY と IME composition の注意点
- **OnEndComposition から直接 write_pty しない**: TSF ロック中に ConPTY に書き込むとタイミング問題が起きる。`PostMessage` で遅延送信する
- **バックスペース（`\x7f`）は1個ずつ送信**: ConPTY は `\x7f` を連続で受け取ると正しく処理しない。`WM_TIMER`（50ms間隔）で1個ずつ送る
- **デバッグログはコミット前に必ず削除**: ファイル I/O を伴うログを高頻度メソッド（GetText 等）に残すと、体感でわかるパフォーマンス悪化を引き起こす

## 技術スタック

| 用途 | クレート/API |
|---|---|
| ConPTY | `windows` crate (`CreatePseudoConsole`, `ReadFile`, `WriteFile`) |
| VT パース + バッファ | `alacritty_terminal` (0.25+) |
| レンダリング | `windows` crate (Direct2D 1.1, DirectWrite) |
| TSF テキストストア | `windows` crate (`ITextStoreACP`, `ITfContextOwnerCompositionSink`, `#[implement]` マクロ) |
| IME 候補位置制御 | `windows` crate (`ImmSetCandidateWindow`, `ImmSetCompositionWindow`) — TSF の `GetTextExt` だけでは不十分 |
| ウィンドウ | `windows` crate (Win32 `CreateWindowExW`, メッセージループ) |

## ITextStoreACP 実装の要点

26 メソッドの実装が必要（大半はスタブで可）。重要なメソッド:

| メソッド | 役割 | 実装方針 |
|---|---|---|
| `GetText` | **核心**。指定範囲のテキストを返す | alacritty_terminal の Grid からカーソル行周辺のテキストを返す |
| `GetSelection` | カーソル位置を返す | ターミナルのカーソル位置を ACP に変換 |
| `GetTextExt` | テキスト範囲の画面座標を返す | IME 候補ウィンドウの位置決めに必要 |
| `GetScreenExt` | ウィンドウの画面座標を返す | ウィンドウ RECT を返す |
| `GetEndACP` | テキスト末尾位置を返す | バッファ全体の文字数 |
| `RequestLock` | ロック管理 | 正しく実装しないと TSF マネージャーがクラッシュ |
| `AdviseSink` | TSF マネージャーのシンク登録 | シンクを保持して変更通知に使う |
| `SetText` | IME がテキストを挿入 | composition 中は preedit バッファに格納、確定時に PTY 転送 |
| `GetStatus` | ドキュメント属性を返す | `TS_SD_READONLY` は返さない（入力可能） |

### IME Composition の実装

TSF の `ITextStoreACP` だけでは IME の composition（変換中テキスト）を正しく扱えない。
以下の追加実装が必要:

| 要素 | 実装 |
|---|---|
| Composition 状態管理 | `ITfContextOwnerCompositionSink` を `ITfSource::AdviseSink` で登録 |
| `OnStartComposition` | composing フラグ ON |
| `OnEndComposition` | composing フラグ OFF、preedit を PTY に送信 |
| preedit バッファ | composition 中は `InsertTextAtSelection`/`SetText` で preedit に格納（PTY 送信しない） |
| 仮想ドキュメント | `GetText`/`GetEndACP` は preedit を含む仮想テキストを返す |
| インライン描画 | preedit テキストをカーソル位置に下線付きで描画 |
| 候補ウィンドウ位置 | **TSF `GetTextExt` + IMM32 `ImmSetCandidateWindow` の併用**が必要（TSF だけでは MS IME の候補位置が反映されない） |

## ConPTY の基本フロー

1. `CreatePipe()` で入出力パイプ作成
2. `CreatePseudoConsole(size, input, output, 0)` で ConPTY 作成
3. `STARTUPINFOEXW` + `PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE` で `wsl.exe` 起動
4. 読み取りスレッドで `ReadFile` → alacritty_terminal でパース → バッファ更新
5. キー入力を `WriteFile` で ConPTY に送信
6. `ResizePseudoConsole` でリサイズ対応

## 重要な罠

### Edition 2024
- `ManuallyDrop` union フィールドへの書き込みに `(*field)` が必要
- `#[unsafe(no_mangle)]` が必要
- `unsafe fn` 本体内でも `unsafe {}` ブロックが必要

### windows crate 0.62
- COM メソッドの引数型は `Ref<'_, T>`（`Option<&T>` ではない）

### WM_CHAR と IME
- `TranslateMessage` は `VK_BACK` を `WM_CHAR(0x08)` に変換する（`0x7F` ではない）。`WM_KEYDOWN` で処理済みのキーは `0x08` と `0x7F` の両方を `WM_CHAR` でフィルタすること
- IME composition 中は `WM_CHAR` と `WM_KEYDOWN` を抑制し、`DefWindowProc` に委譲する（TSF 経由で処理されるため）
- `WM_IME_STARTCOMPOSITION` / `WM_IME_COMPOSITION` で `ImmSetCandidateWindow` を呼び出して候補位置を設定する

### ConPTY
- ReadFile/WriteFile は OVERLAPPED 非対応（同期 I/O のみ）
- 読み取りと書き込みは別スレッドで処理（デッドロック防止）

### レンダリング
- 全角文字は `Flags::WIDE_CHAR` フラグを持ち、次のセルは `WIDE_CHAR_SPACER`。描画時は spacer をスキップし、本体セルで 2 セル幅を描画する

## 類似プロジェクト（参考）

| プロジェクト | 参考ポイント |
|---|---|
| [Alacritty](https://github.com/alacritty/alacritty) | ConPTY 接続、alacritty_terminal の使い方、ターミナルバッファ設計 |
| [WezTerm](https://github.com/wezterm/wezterm) | ターミナルバッファ、Windows 描画 |
| [Windows Terminal](https://github.com/microsoft/terminal) | ConPTY、TSF/IME 統合、Direct3D+DirectWrite 描画 |
| [Chromium TSF](https://chromium.googlesource.com/chromium/src/+/lkgr/ui/base/ime/win/tsf_text_store.cc) | `ITextStoreACP` 実装のリファレンス |
| [Firefox TSFTextStore](https://searchfox.org/mozilla-central/source/widget/windows/TSFTextStore.cpp) | `ITextStoreACP` + composition 処理。レガシー IME 回避策が豊富 |
